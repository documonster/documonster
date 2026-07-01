import { findSignatureFromEnd, hasSignature } from "@archive/__tests__/zip/zip-test-utils";
import { ZipParser } from "@archive/unzip/zip-parser";
import {
  CENTRAL_DIR_HEADER_SIG,
  ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG,
  ZIP64_END_OF_CENTRAL_DIR_SIG,
  END_OF_CENTRAL_DIR_SIG
} from "@archive/zip-spec/zip-records";
import type { ZipEntry } from "@archive/zip/zip-bytes";
import { createZipSync } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

export function runZip64WriteTests(): void {
  describe("ZIP64 write", () => {
    it("createZipSync does not write ZIP64 when not needed", () => {
      const entries: ZipEntry[] = [
        { name: "a.txt", data: new TextEncoder().encode("a") },
        { name: "b.txt", data: new Uint8Array(0) }
      ];
      const zipBytes = createZipSync(entries, { level: 0, reproducible: true });

      expect(
        hasSignature(zipBytes, ZIP64_END_OF_CENTRAL_DIR_SIG, zipBytes.length - 256, zipBytes.length)
      ).toBe(false);
      expect(
        hasSignature(
          zipBytes,
          ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG,
          zipBytes.length - 256,
          zipBytes.length
        )
      ).toBe(false);

      const parser = new ZipParser(zipBytes);
      const parsed = parser.getEntries();
      expect(parsed.length).toBe(2);
    });

    it("createZipSync writes ZIP64 when zip64=true (forced)", () => {
      const entries: ZipEntry[] = [{ name: "a.txt", data: new TextEncoder().encode("a") }];
      const zipBytes = createZipSync(entries, { level: 0, reproducible: true, zip64: true });

      expect(
        hasSignature(zipBytes, ZIP64_END_OF_CENTRAL_DIR_SIG, zipBytes.length - 256, zipBytes.length)
      ).toBe(true);
      expect(
        hasSignature(
          zipBytes,
          ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG,
          zipBytes.length - 256,
          zipBytes.length
        )
      ).toBe(true);

      // Central directory header should use 0xFFFFFFFF sentinels.
      const cdOffset = findSignatureFromEnd(zipBytes, CENTRAL_DIR_HEADER_SIG, 1024 * 1024);
      expect(cdOffset).toBeGreaterThanOrEqual(0);
      const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);
      const compSize32 = view.getUint32(cdOffset + 20, true);
      const uncompSize32 = view.getUint32(cdOffset + 24, true);
      const localOff32 = view.getUint32(cdOffset + 42, true);
      expect(compSize32).toBe(0xffffffff);
      expect(uncompSize32).toBe(0xffffffff);
      expect(localOff32).toBe(0xffffffff);

      const parser = new ZipParser(zipBytes);
      const parsed = parser.getEntries();
      expect(parsed.length).toBe(1);
      expect(parsed[0]!.path).toBe("a.txt");
      expect(parsed[0]!.uncompressedSize).toBe(1);
    });

    it("createZipSync does not write ZIP64 at 65535 entries", () => {
      const count = 65535;
      const entries: ZipEntry[] = new Array(count);
      for (let i = 0; i < count; i++) {
        entries[i] = {
          name: `e/${i}.txt`,
          data: new Uint8Array(0)
        };
      }

      const zipBytes = createZipSync(entries, { level: 0, reproducible: true });

      expect(
        hasSignature(zipBytes, ZIP64_END_OF_CENTRAL_DIR_SIG, zipBytes.length - 256, zipBytes.length)
      ).toBe(false);
      expect(
        hasSignature(
          zipBytes,
          ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG,
          zipBytes.length - 256,
          zipBytes.length
        )
      ).toBe(false);

      const parser = new ZipParser(zipBytes);
      const parsed = parser.getEntries();
      expect(parsed.length).toBe(count);
    }, 30_000);

    it("createZipSync writes ZIP64 EOCD when entry count exceeds 65535", () => {
      const count = 65536;
      const entries: ZipEntry[] = new Array(count);
      for (let i = 0; i < count; i++) {
        entries[i] = {
          name: `e/${i}.txt`,
          data: new Uint8Array(0)
        };
      }

      expect(() =>
        createZipSync(entries, { level: 0, reproducible: true, noSort: true, zip64: false })
      ).toThrow(/ZIP64 is required but zip64=false/);

      // Use noSort to preserve input order, so we can validate specific paths
      const zipBytes = createZipSync(entries, { level: 0, reproducible: true, noSort: true });

      // ZIP64 EOCD + locator should appear near the end.
      expect(
        hasSignature(zipBytes, ZIP64_END_OF_CENTRAL_DIR_SIG, zipBytes.length - 256, zipBytes.length)
      ).toBe(true);
      expect(
        hasSignature(
          zipBytes,
          ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG,
          zipBytes.length - 256,
          zipBytes.length
        )
      ).toBe(true);

      // Validate classic EOCD sentinel fields.
      const eocdOffset = findSignatureFromEnd(zipBytes, END_OF_CENTRAL_DIR_SIG, 1024 * 1024);
      expect(eocdOffset).toBeGreaterThanOrEqual(0);
      const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);
      const entriesOnDisk = view.getUint16(eocdOffset + 8, true);
      const entriesTotal = view.getUint16(eocdOffset + 10, true);
      const cdSize32 = view.getUint32(eocdOffset + 12, true);
      const cdOffset32 = view.getUint32(eocdOffset + 16, true);
      expect(entriesOnDisk).toBe(0xffff);
      expect(entriesTotal).toBe(0xffff);
      expect(cdSize32).toBe(0xffffffff);
      expect(cdOffset32).toBe(0xffffffff);

      // Validate ZIP64 EOCD has correct 64-bit entry counts.
      const zip64EocdOffset = findSignatureFromEnd(
        zipBytes,
        ZIP64_END_OF_CENTRAL_DIR_SIG,
        1024 * 1024
      );
      expect(zip64EocdOffset).toBeGreaterThanOrEqual(0);
      const zip64EntriesOnDisk = view.getBigUint64(zip64EocdOffset + 24, true);
      const zip64EntriesTotal = view.getBigUint64(zip64EocdOffset + 32, true);
      expect(zip64EntriesOnDisk).toBe(BigInt(count));
      expect(zip64EntriesTotal).toBe(BigInt(count));

      const parser = new ZipParser(zipBytes);
      const parsed = parser.getEntries();
      expect(parsed.length).toBe(count);
      expect(parsed[0]!.path).toBe("e/0.txt");
      expect(parsed[count - 1]!.path).toBe(`e/${count - 1}.txt`);
    }, 30_000);
  });
}
