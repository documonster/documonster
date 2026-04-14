/**
 * Tests for ZIP timestamp utilities.
 */

import {
  dateToDos,
  parseDosDateTimeUTC,
  dateToZipDos,
  buildZipTimestampExtraField,
  resolveZipLastModifiedDateFromUnixSeconds,
  EXTENDED_TIMESTAMP_ID,
  NTFS_TIMESTAMP_ID
} from "@archive/zip-spec/timestamps";
import { describe, it, expect } from "vitest";

describe("timestamps", () => {
  describe("dateToDos", () => {
    it("should convert a date to DOS time and date fields", () => {
      // 2024-06-15 14:30:20 local time
      const date = new Date(2024, 5, 15, 14, 30, 20);
      const [dosTime, dosDate] = dateToDos(date);

      // DOS time: hours(5bits) << 11 | minutes(6bits) << 5 | seconds/2(5bits)
      // 14 << 11 | 30 << 5 | 10 = 0x7390 + 0x3C0 + 0xA = 0x773A
      expect(dosTime).toBe((14 << 11) | (30 << 5) | 10);

      // DOS date: (year-1980)(7bits) << 9 | month(4bits) << 5 | day(5bits)
      // (2024-1980) << 9 | 6 << 5 | 15 = 44 << 9 | 192 | 15 = 0x58CF
      expect(dosDate).toBe((44 << 9) | (6 << 5) | 15);
    });

    it("should handle year 1980 (minimum)", () => {
      const date = new Date(1980, 0, 1, 0, 0, 0);
      const [dosTime, dosDate] = dateToDos(date);

      expect(dosTime).toBe(0);
      expect(dosDate).toBe((0 << 9) | (1 << 5) | 1);
    });

    it("should handle seconds with 2-second resolution", () => {
      // DOS time has 2-second resolution, so 21 seconds becomes 10
      const date = new Date(2024, 0, 1, 12, 0, 21);
      const [dosTime] = dateToDos(date);

      expect(dosTime & 0x1f).toBe(10); // 21 >> 1 = 10
    });
  });

  describe("parseDosDateTimeUTC", () => {
    it("should parse DOS date without time", () => {
      // DOS date for 2024-06-15
      const dosDate = (44 << 9) | (6 << 5) | 15;
      const result = parseDosDateTimeUTC(dosDate);

      expect(result.getUTCFullYear()).toBe(2024);
      expect(result.getUTCMonth()).toBe(5); // June (0-indexed)
      expect(result.getUTCDate()).toBe(15);
      expect(result.getUTCHours()).toBe(0);
      expect(result.getUTCMinutes()).toBe(0);
      expect(result.getUTCSeconds()).toBe(0);
    });

    it("should parse DOS date and time", () => {
      // DOS date for 2024-06-15, time for 14:30:20
      const dosDate = (44 << 9) | (6 << 5) | 15;
      const dosTime = (14 << 11) | (30 << 5) | 10;
      const result = parseDosDateTimeUTC(dosDate, dosTime);

      expect(result.getUTCFullYear()).toBe(2024);
      expect(result.getUTCMonth()).toBe(5);
      expect(result.getUTCDate()).toBe(15);
      expect(result.getUTCHours()).toBe(14);
      expect(result.getUTCMinutes()).toBe(30);
      expect(result.getUTCSeconds()).toBe(20); // 10 * 2 = 20
    });

    it("should handle year 1980 (base year)", () => {
      const dosDate = (0 << 9) | (1 << 5) | 1; // 1980-01-01
      const result = parseDosDateTimeUTC(dosDate);

      expect(result.getUTCFullYear()).toBe(1980);
      expect(result.getUTCMonth()).toBe(0);
      expect(result.getUTCDate()).toBe(1);
    });

    it("should handle year 2107 (maximum DOS year)", () => {
      const dosDate = (127 << 9) | (12 << 5) | 31; // 2107-12-31
      const result = parseDosDateTimeUTC(dosDate);

      expect(result.getUTCFullYear()).toBe(2107);
      expect(result.getUTCMonth()).toBe(11);
      expect(result.getUTCDate()).toBe(31);
    });
  });

  describe("dateToZipDos", () => {
    it("should return object with dosTime and dosDate", () => {
      const date = new Date(2024, 5, 15, 14, 30, 20);
      const result = dateToZipDos(date);

      expect(result).toHaveProperty("dosTime");
      expect(result).toHaveProperty("dosDate");

      const [dosTime, dosDate] = dateToDos(date);
      expect(result.dosTime).toBe(dosTime);
      expect(result.dosDate).toBe(dosDate);
    });
  });

  describe("buildZipTimestampExtraField", () => {
    const testDate = new Date(Date.UTC(2024, 5, 15, 12, 0, 0));

    it("should return empty array for 'dos' mode", () => {
      const result = buildZipTimestampExtraField(testDate, "dos");
      expect(result.length).toBe(0);
    });

    it("should build extended timestamp for 'dos+utc' mode", () => {
      const result = buildZipTimestampExtraField(testDate, "dos+utc");

      // Should have: header(4) + flags(1) + mtime(4) = 9 bytes
      expect(result.length).toBe(9);

      const view = new DataView(result.buffer, result.byteOffset, result.byteLength);
      expect(view.getUint16(0, true)).toBe(EXTENDED_TIMESTAMP_ID);
      expect(view.getUint16(2, true)).toBe(5); // data size: 1 + 4

      // flags: bit 0 = mtime present
      expect(result[4]).toBe(0x01);

      // mtime as Unix seconds
      const unixSeconds = Math.floor(testDate.getTime() / 1000);
      expect(view.getUint32(5, true)).toBe(unixSeconds);
    });

    it("should include atime and ctime when provided", () => {
      const atime = new Date(Date.UTC(2024, 5, 14, 10, 0, 0));
      const ctime = new Date(Date.UTC(2024, 5, 13, 8, 0, 0));

      const result = buildZipTimestampExtraField(testDate, "dos+utc", { atime, ctime });

      // header(4) + flags(1) + mtime(4) + atime(4) + ctime(4) = 17 bytes
      expect(result.length).toBe(17);

      const view = new DataView(result.buffer, result.byteOffset, result.byteLength);

      // flags: mtime(1) + atime(2) + ctime(4) = 7
      expect(result[4]).toBe(0x07);

      // Check timestamps
      expect(view.getUint32(5, true)).toBe(Math.floor(testDate.getTime() / 1000));
      expect(view.getUint32(9, true)).toBe(Math.floor(atime.getTime() / 1000));
      expect(view.getUint32(13, true)).toBe(Math.floor(ctime.getTime() / 1000));
    });

    it("should build both extended and NTFS timestamps for 'dos+utc+ntfs' mode", () => {
      const result = buildZipTimestampExtraField(testDate, "dos+utc+ntfs");

      // Extended: 9 bytes
      // NTFS: header(4) + reserved(4) + tag(2) + size(2) + 4*filetime(32) = 44 bytes
      // Total: 53 bytes
      expect(result.length).toBe(9 + 44);

      const view = new DataView(result.buffer, result.byteOffset, result.byteLength);

      // First field: Extended timestamp
      expect(view.getUint16(0, true)).toBe(EXTENDED_TIMESTAMP_ID);

      // Second field: NTFS timestamp
      expect(view.getUint16(9, true)).toBe(NTFS_TIMESTAMP_ID);
      expect(view.getUint16(11, true)).toBe(40); // data size

      // NTFS tag 0x0001, size 32
      expect(view.getUint16(17, true)).toBe(0x0001);
      expect(view.getUint16(19, true)).toBe(32);
    });

    it("should include birthTime in NTFS field", () => {
      const birthTime = new Date(Date.UTC(2024, 5, 10, 6, 0, 0));

      const result = buildZipTimestampExtraField(testDate, "dos+utc+ntfs", { birthTime });

      // NTFS field starts at offset 9
      const view = new DataView(result.buffer, result.byteOffset, result.byteLength);

      // FILETIME for birthTime is at offset 9 + 12 + 24 = 45
      const filetime = view.getBigUint64(45, true);

      // Verify it's a valid FILETIME (100-nanosecond intervals since 1601)
      // For 2024, it should be > 0
      expect(filetime).toBeGreaterThan(0n);
    });
  });

  describe("resolveZipLastModifiedDateFromUnixSeconds", () => {
    it("should use Unix seconds when provided", () => {
      const unixSeconds = Math.floor(Date.UTC(2024, 5, 15, 12, 0, 0) / 1000);
      const result = resolveZipLastModifiedDateFromUnixSeconds(0, 0, unixSeconds);

      expect(result.getTime()).toBe(unixSeconds * 1000);
    });

    it("should fall back to DOS date/time when Unix seconds is undefined", () => {
      const dosDate = (44 << 9) | (6 << 5) | 15; // 2024-06-15
      const dosTime = (12 << 11) | (30 << 5) | 10; // 12:30:20

      const result = resolveZipLastModifiedDateFromUnixSeconds(dosDate, dosTime, undefined);

      expect(result.getUTCFullYear()).toBe(2024);
      expect(result.getUTCMonth()).toBe(5);
      expect(result.getUTCDate()).toBe(15);
      expect(result.getUTCHours()).toBe(12);
      expect(result.getUTCMinutes()).toBe(30);
    });
  });
});
