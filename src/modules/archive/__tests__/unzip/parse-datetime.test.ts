import { parseDosDateTime } from "@archive/zip-spec/timestamps";
import { describe, it, expect } from "vitest";

describe("parse-datetime", () => {
  describe("parseDosDateTime", () => {
    it("should parse DOS date without time", () => {
      // Date: 2020-06-15 (year=40 since 1980, month=6, day=15)
      // Binary: year(7bits)=0101000, month(4bits)=0110, day(5bits)=01111
      // Combined: 0101000 0110 01111 = 0x5 0CF = 0x50CF
      const date = (40 << 9) | (6 << 5) | 15; // 0x50CF
      const result = parseDosDateTime(date);
      expect(result.getFullYear()).toBe(2020);
      expect(result.getMonth()).toBe(5); // June (0-indexed)
      expect(result.getDate()).toBe(15);
      expect(result.getHours()).toBe(0);
      expect(result.getMinutes()).toBe(0);
      expect(result.getSeconds()).toBe(0);
    });

    it("should parse DOS date and time", () => {
      // Date: 2020-06-15
      const date = (40 << 9) | (6 << 5) | 15;
      // Time: 14:30:22 (hours=14, minutes=30, seconds=22/2=11)
      // Binary: hours(5bits)=01110, minutes(6bits)=011110, seconds(5bits)=01011
      // Combined: 01110 011110 01011 = 0x73CB
      const time = (14 << 11) | (30 << 5) | 11;
      const result = parseDosDateTime(date, time);
      expect(result.getFullYear()).toBe(2020);
      expect(result.getMonth()).toBe(5);
      expect(result.getDate()).toBe(15);
      expect(result.getHours()).toBe(14);
      expect(result.getMinutes()).toBe(30);
      expect(result.getSeconds()).toBe(22);
    });

    it("should handle minimum date (1980-01-01)", () => {
      const date = (0 << 9) | (1 << 5) | 1;
      const result = parseDosDateTime(date);
      expect(result.getFullYear()).toBe(1980);
      expect(result.getMonth()).toBe(0);
      expect(result.getDate()).toBe(1);
    });

    it("should handle time with seconds", () => {
      const date = (40 << 9) | (1 << 5) | 1;
      // Time: 00:00:58 (seconds stored as 29)
      const time = (0 << 11) | (0 << 5) | 29;
      const result = parseDosDateTime(date, time);
      expect(result.getHours()).toBe(0);
      expect(result.getMinutes()).toBe(0);
      expect(result.getSeconds()).toBe(58);
    });

    it("should handle end of day time", () => {
      const date = (40 << 9) | (1 << 5) | 1;
      // Time: 23:59:58
      const time = (23 << 11) | (59 << 5) | 29;
      const result = parseDosDateTime(date, time);
      expect(result.getHours()).toBe(23);
      expect(result.getMinutes()).toBe(59);
      expect(result.getSeconds()).toBe(58);
    });
  });
});
