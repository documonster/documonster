/**
 * End-to-end tests for ZIP encryption (createZip/createZipSync + ZipParser)
 *
 * These tests cover edge cases not handled by the streaming-zip tests:
 * - Buffer-based API (createZip/createZipSync) with encryption
 * - Wrong password handling at ZipParser level
 * - Mixed encryption (some files encrypted, some not)
 * - Empty file encryption
 * - AES-192 key strength
 */

import { describe, it, expect } from "vitest";
import { createZip, createZipSync, type ZipEntry } from "@archive/zip/zip-bytes";
import { ZipParser } from "@archive/unzip/zip-parser";

describe("ZIP Encryption End-to-End", () => {
  describe("createZip with ZipCrypto", () => {
    it("should create and extract ZipCrypto-encrypted ZIP", async () => {
      const entries: ZipEntry[] = [
        {
          name: "secret.txt",
          data: new TextEncoder().encode("This is secret content")
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "my-password"
      });

      // Parse and verify encryption flag
      const parser = new ZipParser(zipData);
      const entryList = parser.getEntries();
      expect(entryList.length).toBe(1);
      expect(entryList[0]!.isEncrypted).toBe(true);
      expect(entryList[0]!.encryptionMethod).toBe("zipcrypto");

      // Extract with correct password
      const extracted = await parser.extractAll("my-password");
      const content = new TextDecoder().decode(extracted.get("secret.txt")!);
      expect(content).toBe("This is secret content");
    });

    it("should fail with wrong password (ZipCrypto)", async () => {
      const entries: ZipEntry[] = [
        { name: "data.txt", data: new TextEncoder().encode("Test data") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "correct"
      });

      const parser = new ZipParser(zipData);
      // ZipCrypto header check has 1/256 false positive rate per ZIP spec.
      // Wrong password always fails, but error may be DecryptionError or CRC32/inflate mismatch.
      await expect(parser.extract("data.txt", "wrong")).rejects.toThrow(
        /incorrect password|CRC32 mismatch|size mismatch|invalid/i
      );
    });

    it("should throw when encryption without password (createZip)", async () => {
      const entries: ZipEntry[] = [{ name: "test.txt", data: new Uint8Array(10) }];

      await expect(createZip(entries, { encryptionMethod: "zipcrypto" })).rejects.toThrow(
        /password is required/i
      );
    });
  });

  describe("createZipSync with ZipCrypto", () => {
    it("should create and extract ZipCrypto-encrypted ZIP (sync)", async () => {
      const entries: ZipEntry[] = [
        {
          name: "sync-secret.txt",
          data: new TextEncoder().encode("Sync secret")
        }
      ];

      const zipData = createZipSync(entries, {
        encryptionMethod: "zipcrypto",
        password: "sync-password"
      });

      const parser = new ZipParser(zipData);
      expect(parser.getEntries()[0]!.isEncrypted).toBe(true);

      // ZipCrypto can be extracted synchronously
      const extracted = parser.extractSync("sync-secret.txt", "sync-password");
      expect(new TextDecoder().decode(extracted!)).toBe("Sync secret");
    });

    it("should throw when using AES with createZipSync", () => {
      const entries: ZipEntry[] = [{ name: "test.txt", data: new Uint8Array(10) }];

      expect(() =>
        createZipSync(entries, {
          encryptionMethod: "aes-256",
          password: "test"
        })
      ).toThrow(/async API/i);
    });
  });

  describe("createZip with AES", () => {
    it("should create and extract AES-256 encrypted ZIP", async () => {
      const entries: ZipEntry[] = [
        {
          name: "aes-file.txt",
          data: new TextEncoder().encode("AES-256 protected")
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "strong-password"
      });

      const parser = new ZipParser(zipData);
      const entryList = parser.getEntries();
      expect(entryList[0]!.isEncrypted).toBe(true);
      expect(entryList[0]!.encryptionMethod).toBe("aes");
      expect(entryList[0]!.aesKeyStrength).toBe(256);

      const extracted = await parser.extractAll("strong-password");
      const content = new TextDecoder().decode(extracted.get("aes-file.txt")!);
      expect(content).toBe("AES-256 protected");
    });

    it("should create and extract AES-192 encrypted ZIP", async () => {
      const entries: ZipEntry[] = [
        {
          name: "aes192.txt",
          data: new TextEncoder().encode("AES-192 content")
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-192",
        password: "aes192-password"
      });

      const parser = new ZipParser(zipData);
      const entryList = parser.getEntries();
      expect(entryList[0]!.aesKeyStrength).toBe(192);

      const extracted = await parser.extractAll("aes192-password");
      expect(new TextDecoder().decode(extracted.get("aes192.txt")!)).toBe("AES-192 content");
    });

    it("should create and extract AES-128 encrypted ZIP", async () => {
      const entries: ZipEntry[] = [
        {
          name: "aes128.txt",
          data: new TextEncoder().encode("AES-128 content")
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-128",
        password: "aes128-password"
      });

      const parser = new ZipParser(zipData);
      const entryList = parser.getEntries();
      expect(entryList[0]!.aesKeyStrength).toBe(128);

      const extracted = await parser.extractAll("aes128-password");
      expect(new TextDecoder().decode(extracted.get("aes128.txt")!)).toBe("AES-128 content");
    });

    it("should fail with wrong password (AES)", async () => {
      const entries: ZipEntry[] = [
        { name: "data.txt", data: new TextEncoder().encode("Test data") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "correct"
      });

      const parser = new ZipParser(zipData);
      await expect(parser.extract("data.txt", "wrong")).rejects.toThrow(
        /password|verification failed/i
      );
    });

    it("should throw when extracting AES with extractSync", async () => {
      const entries: ZipEntry[] = [
        { name: "aes.txt", data: new TextEncoder().encode("AES content") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      expect(() => parser.extractSync("aes.txt", "password")).toThrow(/async/i);
    });
  });

  describe("Mixed encryption", () => {
    it("should handle per-entry encryption override", async () => {
      const entries: ZipEntry[] = [
        {
          name: "encrypted.txt",
          data: new TextEncoder().encode("Encrypted"),
          encryptionMethod: "zipcrypto",
          password: "file-password"
        },
        {
          name: "plain.txt",
          data: new TextEncoder().encode("Not encrypted"),
          encryptionMethod: "none"
        }
      ];

      const zipData = await createZip(entries);

      const parser = new ZipParser(zipData);
      const entryList = parser.getEntries();

      const encrypted = entryList.find(e => e.path === "encrypted.txt");
      const plain = entryList.find(e => e.path === "plain.txt");

      expect(encrypted!.isEncrypted).toBe(true);
      expect(plain!.isEncrypted).toBe(false);

      // Extract encrypted file with password
      const encryptedContent = await parser.extract("encrypted.txt", "file-password");
      expect(new TextDecoder().decode(encryptedContent!)).toBe("Encrypted");

      // Extract plain file without password
      const plainContent = await parser.extract("plain.txt");
      expect(new TextDecoder().decode(plainContent!)).toBe("Not encrypted");
    });

    it("should use archive password as default", async () => {
      const entries: ZipEntry[] = [
        { name: "file1.txt", data: new TextEncoder().encode("File 1") },
        { name: "file2.txt", data: new TextEncoder().encode("File 2") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "archive-password"
      });

      const parser = new ZipParser(zipData, { password: "archive-password" });

      // Should extract without explicit password since constructor password is set
      const content1 = await parser.extract("file1.txt");
      const content2 = await parser.extract("file2.txt");

      expect(new TextDecoder().decode(content1!)).toBe("File 1");
      expect(new TextDecoder().decode(content2!)).toBe("File 2");
    });
  });

  describe("Edge cases", () => {
    it("should encrypt empty files", async () => {
      const entries: ZipEntry[] = [{ name: "empty.txt", data: new Uint8Array(0) }];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      expect(parser.getEntries()[0]!.isEncrypted).toBe(true);

      const extracted = await parser.extract("empty.txt", "password");
      expect(extracted!.length).toBe(0);
    });

    it("should encrypt empty files with AES", async () => {
      const entries: ZipEntry[] = [{ name: "empty-aes.txt", data: new Uint8Array(0) }];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      expect(parser.getEntries()[0]!.isEncrypted).toBe(true);
      expect(parser.getEntries()[0]!.encryptionMethod).toBe("aes");

      const extracted = await parser.extract("empty-aes.txt", "password");
      expect(extracted!.length).toBe(0);
    });

    it("should handle large files with encryption", async () => {
      // 100KB of random data
      const largeData = new Uint8Array(100 * 1024);
      for (let i = 0; i < largeData.length; i++) {
        largeData[i] = i % 256;
      }

      const entries: ZipEntry[] = [{ name: "large.bin", data: largeData }];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      const extracted = await parser.extract("large.bin", "password");

      expect(extracted).toEqual(largeData);
    });

    it("should handle STORE mode (level=0) with encryption", async () => {
      const entries: ZipEntry[] = [
        {
          name: "stored.txt",
          data: new TextEncoder().encode("Stored content"),
          level: 0
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "password",
        level: 0
      });

      const parser = new ZipParser(zipData);
      const entry = parser.getEntries()[0]!;
      expect(entry.isEncrypted).toBe(true);
      // For ZipCrypto, compression method should still be STORE
      expect(entry.compressionMethod).toBe(0);

      const extracted = await parser.extract("stored.txt", "password");
      expect(new TextDecoder().decode(extracted!)).toBe("Stored content");
    });

    it("should handle STORE mode with AES encryption", async () => {
      const entries: ZipEntry[] = [
        {
          name: "aes-stored.txt",
          data: new TextEncoder().encode("AES Stored"),
          level: 0
        }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "aes-256",
        password: "password",
        level: 0
      });

      const parser = new ZipParser(zipData);
      const entry = parser.getEntries()[0]!;
      expect(entry.isEncrypted).toBe(true);
      // For AES, compression method is 99 (AES indicator)
      expect(entry.compressionMethod).toBe(99);
      // Original compression method should be STORE
      expect(entry.originalCompressionMethod).toBe(0);

      const extracted = await parser.extract("aes-stored.txt", "password");
      expect(new TextDecoder().decode(extracted!)).toBe("AES Stored");
    });

    it("should throw when extracting encrypted file without password", async () => {
      const entries: ZipEntry[] = [
        { name: "encrypted.txt", data: new TextEncoder().encode("Secret") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      await expect(parser.extract("encrypted.txt")).rejects.toThrow(/password/i);
    });
  });

  describe("ZipParser helper methods", () => {
    it("should detect encrypted entries with hasEncryptedEntries", async () => {
      const entries: ZipEntry[] = [
        { name: "encrypted.txt", data: new TextEncoder().encode("Encrypted") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "password"
      });

      const parser = new ZipParser(zipData);
      expect(parser.hasEncryptedEntries()).toBe(true);
    });

    it("should return false for hasEncryptedEntries on plain ZIP", async () => {
      const entries: ZipEntry[] = [{ name: "plain.txt", data: new TextEncoder().encode("Plain") }];

      const zipData = await createZip(entries);

      const parser = new ZipParser(zipData);
      expect(parser.hasEncryptedEntries()).toBe(false);
    });

    it("should list encrypted entries with getEncryptedEntries", async () => {
      const entries: ZipEntry[] = [
        {
          name: "encrypted.txt",
          data: new TextEncoder().encode("Encrypted"),
          encryptionMethod: "zipcrypto",
          password: "password"
        },
        { name: "plain.txt", data: new TextEncoder().encode("Plain") }
      ];

      const zipData = await createZip(entries);

      const parser = new ZipParser(zipData);
      const encryptedEntries = parser.getEncryptedEntries();

      expect(encryptedEntries.length).toBe(1);
      expect(encryptedEntries[0]!.path).toBe("encrypted.txt");
    });

    it("should allow setting password after construction", async () => {
      const entries: ZipEntry[] = [
        { name: "data.txt", data: new TextEncoder().encode("Secret data") }
      ];

      const zipData = await createZip(entries, {
        encryptionMethod: "zipcrypto",
        password: "password"
      });

      const parser = new ZipParser(zipData);

      // Should fail without password
      await expect(parser.extract("data.txt")).rejects.toThrow();

      // Set password
      parser.setPassword("password");

      // Should work now
      const extracted = await parser.extract("data.txt");
      expect(new TextDecoder().decode(extracted!)).toBe("Secret data");
    });
  });

  describe("UnzipEntry.isEncrypted", () => {
    it("should be true for encrypted entries in buffer mode", async () => {
      const zipData = await createZip(
        [{ name: "secret.txt", data: new TextEncoder().encode("data") }],
        { encryptionMethod: "zipcrypto", password: "pw" }
      );

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(zipData, { password: "pw" });
      const entries: any[] = [];
      for await (const entry of reader.entries()) {
        entries.push(entry);
        await entry.bytes(); // consume
      }
      expect(entries).toHaveLength(1);
      expect(entries[0].isEncrypted).toBe(true);
    });

    it("should be false for non-encrypted entries in buffer mode", async () => {
      const zipData = await createZip(
        [{ name: "plain.txt", data: new TextEncoder().encode("data") }],
        { level: 0, smartStore: false }
      );

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(zipData);
      const entries: any[] = [];
      for await (const entry of reader.entries()) {
        entries.push(entry);
        await entry.bytes();
      }
      expect(entries).toHaveLength(1);
      expect(entries[0].isEncrypted).toBe(false);
    });

    it("should be true for encrypted entries in streaming mode", async () => {
      const zipData = await createZip(
        [{ name: "secret.txt", data: new TextEncoder().encode("data") }],
        { encryptionMethod: "zipcrypto", password: "pw" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);
      const entries: any[] = [];
      for await (const entry of reader.entries()) {
        entries.push(entry);
        entry.discard();
      }
      expect(entries).toHaveLength(1);
      expect(entries[0].isEncrypted).toBe(true);
    });

    it("should be false for non-encrypted entries in streaming mode", async () => {
      const zipData = await createZip(
        [{ name: "plain.txt", data: new TextEncoder().encode("data") }],
        { level: 0, smartStore: false }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);
      const entries: any[] = [];
      for await (const entry of reader.entries()) {
        entries.push(entry);
        entry.discard();
      }
      expect(entries).toHaveLength(1);
      expect(entries[0].isEncrypted).toBe(false);
    });

    it("should not throw inflate error for ZipCrypto-encrypted entry in streaming mode", async () => {
      const zipData = await createZip(
        [
          {
            name: "secret.txt",
            data: new TextEncoder().encode(
              "This is secret content that is encrypted with a password"
            )
          }
        ],
        { encryptionMethod: "zipcrypto", password: "my-password" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);

      let validCount = 0;
      let firstPath = "";
      let firstEncrypted = false;

      for await (const entry of reader.entries()) {
        validCount++;
        if (validCount === 1) {
          firstPath = entry.path;
          firstEncrypted = entry.isEncrypted;
        } else {
          entry.discard();
          break;
        }
        entry.discard();
      }

      expect(validCount).toBe(1);
      expect(firstPath).toBe("secret.txt");
      expect(firstEncrypted).toBe(true);
    });

    it("should not throw inflate error for AES-encrypted entry in streaming mode", async () => {
      const zipData = await createZip(
        [
          {
            name: "aes-secret.txt",
            data: new TextEncoder().encode("AES encrypted content for testing")
          }
        ],
        { encryptionMethod: "aes-256", password: "aes-password" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);

      const entries: any[] = [];
      for await (const entry of reader.entries()) {
        entries.push({ path: entry.path, isEncrypted: entry.isEncrypted });
        entry.discard();
      }

      expect(entries).toHaveLength(1);
      expect(entries[0].path).toBe("aes-secret.txt");
      expect(entries[0].isEncrypted).toBe(true);
    });

    it("should decrypt ZipCrypto entry via bytes() in streaming mode", async () => {
      const content = "This is secret content that is encrypted with a password";
      const zipData = await createZip(
        [{ name: "secret.txt", data: new TextEncoder().encode(content) }],
        { encryptionMethod: "zipcrypto", password: "my-password" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream, { password: "my-password" });

      for await (const entry of reader.entries()) {
        expect(entry.isEncrypted).toBe(true);
        const bytes = await entry.bytes();
        const text = new TextDecoder().decode(bytes);
        expect(text).toBe(content);
      }
    });

    it("should decrypt AES-256 entry via bytes() in streaming mode", async () => {
      const content = "AES encrypted content for streaming decryption test";
      const zipData = await createZip(
        [{ name: "aes-secret.txt", data: new TextEncoder().encode(content) }],
        { encryptionMethod: "aes-256", password: "aes-password" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream, { password: "aes-password" });

      for await (const entry of reader.entries()) {
        expect(entry.isEncrypted).toBe(true);
        const bytes = await entry.bytes();
        const text = new TextDecoder().decode(bytes);
        expect(text).toBe(content);
      }
    });

    it("should throw PasswordRequiredError for encrypted entry without password in streaming mode", async () => {
      const zipData = await createZip(
        [{ name: "secret.txt", data: new TextEncoder().encode("secret") }],
        { encryptionMethod: "zipcrypto", password: "pw" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream); // no password

      for await (const entry of reader.entries()) {
        expect(entry.isEncrypted).toBe(true);
        await expect(entry.bytes()).rejects.toThrow(/password/i);
      }
    });
  });

  describe("discard() hang regression", () => {
    it("should not hang when discarding large encrypted entries (zipcrypto)", async () => {
      const largeData = new Uint8Array(512 * 1024).fill(65);

      const zipData = await createZip(
        [
          { name: "large-secret.txt", data: largeData },
          { name: "another.txt", data: new TextEncoder().encode("hello") }
        ],
        { encryptionMethod: "zipcrypto", password: "pw" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);
      const entries: { path: string; isEncrypted: boolean }[] = [];

      for await (const entry of reader.entries()) {
        entries.push({ path: entry.path, isEncrypted: entry.isEncrypted });
        entry.discard();
      }

      expect(entries.length).toBe(2);
      expect(entries[0].isEncrypted).toBe(true);
      expect(entries[1].isEncrypted).toBe(true);
    }, 5000);

    it("should not hang when discarding large encrypted entries (aes-256)", async () => {
      const largeData = new Uint8Array(512 * 1024).fill(66);

      const zipData = await createZip(
        [
          { name: "large-aes.txt", data: largeData },
          { name: "small.txt", data: new TextEncoder().encode("world") }
        ],
        { encryptionMethod: "aes-256", password: "pw" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { ZipReader } = await import("@archive/unzip");
      const reader = new ZipReader(stream);
      const entries: { path: string; isEncrypted: boolean }[] = [];

      for await (const entry of reader.entries()) {
        entries.push({ path: entry.path, isEncrypted: entry.isEncrypted });
        entry.discard();
      }

      expect(entries.length).toBe(2);
      expect(entries[0].isEncrypted).toBe(true);
      expect(entries[1].isEncrypted).toBe(true);
    }, 5000);

    it("should not hang with unzip() high-level API", async () => {
      const largeData = new Uint8Array(1024 * 1024).fill(67); // 1MB

      const zipData = await createZip(
        [
          { name: "big.txt", data: largeData },
          { name: "tiny.txt", data: new TextEncoder().encode("x") }
        ],
        { encryptionMethod: "zipcrypto", password: "pw" }
      );

      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(zipData);
          controller.close();
        }
      });

      const { unzip } = await import("@archive/read-archive");
      const reader = unzip(stream);
      const entries: { path: string; isEncrypted: boolean }[] = [];

      for await (const entry of reader.entries()) {
        entries.push({ path: entry.path, isEncrypted: entry.isEncrypted });
        entry.discard();
      }

      expect(entries.length).toBe(2);
      expect(entries[0].isEncrypted).toBe(true);
    }, 5000);
  });
});
