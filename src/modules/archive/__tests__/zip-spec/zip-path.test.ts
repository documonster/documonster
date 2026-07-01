/**
 * Tests for ZIP path normalization and joining utilities.
 */

import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import { normalizeZipPath, joinZipPath } from "@archive/zip-spec/zip-path";
import { describe, it, expect } from "vitest";

describe("zip-path", () => {
  describe("normalizeZipPath", () => {
    describe("legacy mode (default)", () => {
      it("should replace backslashes with forward slashes", () => {
        expect(normalizeZipPath("foo\\bar\\baz.txt")).toBe("foo/bar/baz.txt");
        expect(normalizeZipPath("a\\b\\c\\d")).toBe("a/b/c/d");
      });

      it("should strip leading slashes", () => {
        expect(normalizeZipPath("/foo/bar.txt")).toBe("foo/bar.txt");
        expect(normalizeZipPath("///foo/bar.txt")).toBe("foo/bar.txt");
      });

      it("should handle mixed slashes", () => {
        expect(normalizeZipPath("/foo\\bar/baz\\qux")).toBe("foo/bar/baz/qux");
      });

      it("should not normalize . and .. segments in legacy mode", () => {
        expect(normalizeZipPath("foo/./bar")).toBe("foo/./bar");
        expect(normalizeZipPath("foo/../bar")).toBe("foo/../bar");
      });

      it("should not strip Windows drive in legacy mode", () => {
        expect(normalizeZipPath("C:\\foo\\bar")).toBe("C:/foo/bar");
        expect(normalizeZipPath("D:/foo/bar")).toBe("D:/foo/bar");
      });

      it("should handle empty string", () => {
        expect(normalizeZipPath("")).toBe("");
      });

      it("should prepend slash when option is set", () => {
        expect(normalizeZipPath("foo/bar.txt", { prependSlash: true })).toBe("/foo/bar.txt");
        expect(normalizeZipPath("/foo/bar.txt", { prependSlash: true })).toBe("/foo/bar.txt");
      });
    });

    describe("posix mode", () => {
      const opts: ZipPathOptions = { mode: "posix" };

      it("should normalize . segments", () => {
        expect(normalizeZipPath("foo/./bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("./foo/bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("foo/bar/.", opts)).toBe("foo/bar");
      });

      it("should normalize .. segments", () => {
        expect(normalizeZipPath("foo/bar/../baz", opts)).toBe("foo/baz");
        expect(normalizeZipPath("foo/bar/baz/../../qux", opts)).toBe("foo/qux");
      });

      it("should preserve single leading .. but collapse multiple", () => {
        // Single leading ".." is preserved in posix mode
        expect(normalizeZipPath("../foo", opts)).toBe("../foo");
        // Multiple leading ".." collapses: ../../foo -> foo (pops "..", leaves "foo")
        expect(normalizeZipPath("../../foo", opts)).toBe("foo");
        // Normal parent traversal within path
        expect(normalizeZipPath("foo/../bar", opts)).toBe("bar");
      });

      it("should strip Windows drive by default", () => {
        expect(normalizeZipPath("C:\\foo\\bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("D:/foo/bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("C:foo", opts)).toBe("foo");
      });

      it("should preserve Windows drive when stripDrive is false", () => {
        expect(normalizeZipPath("C:/foo/bar", { mode: "posix", stripDrive: false })).toBe(
          "C:/foo/bar"
        );
      });

      it("should strip leading slashes", () => {
        expect(normalizeZipPath("/foo/bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("///foo/bar", opts)).toBe("foo/bar");
      });

      it("should collapse multiple slashes", () => {
        expect(normalizeZipPath("foo//bar///baz", opts)).toBe("foo/bar/baz");
      });

      it("should handle complex path", () => {
        expect(normalizeZipPath("C:\\Users\\./foo/../bar/./baz", opts)).toBe("Users/bar/baz");
      });

      it("should handle empty string", () => {
        expect(normalizeZipPath("", opts)).toBe("");
      });

      it("should prepend slash when option is set", () => {
        expect(normalizeZipPath("foo/bar", { mode: "posix", prependSlash: true })).toBe("/foo/bar");
      });
    });

    describe("safe mode", () => {
      const opts: ZipPathOptions = { mode: "safe" };

      it("should reject absolute paths", () => {
        expect(() => normalizeZipPath("/foo/bar", opts)).toThrow("Unsafe ZIP path (absolute)");
        expect(() => normalizeZipPath("///foo", opts)).toThrow("Unsafe ZIP path (absolute)");
      });

      it("should reject paths with .. traversal", () => {
        expect(() => normalizeZipPath("../foo", opts)).toThrow("Unsafe ZIP path (traversal)");
        expect(() => normalizeZipPath("foo/../../bar", opts)).toThrow(
          "Unsafe ZIP path (traversal)"
        );
      });

      it("should reject Windows drive when stripDrive is false", () => {
        expect(() => normalizeZipPath("C:/foo", { mode: "safe", stripDrive: false })).toThrow(
          "Unsafe ZIP path (drive)"
        );
      });

      it("should allow safe relative paths", () => {
        expect(normalizeZipPath("foo/bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("foo/./bar", opts)).toBe("foo/bar");
        expect(normalizeZipPath("foo/bar/../baz", opts)).toBe("foo/baz");
      });

      it("should strip Windows drive by default", () => {
        expect(normalizeZipPath("C:\\foo\\bar", opts)).toBe("foo/bar");
      });

      it("should allow just ..", () => {
        expect(() => normalizeZipPath("..", opts)).toThrow("Unsafe ZIP path (traversal)");
      });

      it("should allow empty path after normalization", () => {
        expect(normalizeZipPath("", opts)).toBe("");
        expect(normalizeZipPath(".", opts)).toBe("");
        expect(normalizeZipPath("foo/..", opts)).toBe("");
      });

      it("should prepend slash when option is set", () => {
        expect(normalizeZipPath("foo/bar", { mode: "safe", prependSlash: true })).toBe("/foo/bar");
      });
    });
  });

  describe("joinZipPath", () => {
    it("should join multiple path segments", () => {
      expect(joinZipPath({}, "foo", "bar", "baz.txt")).toBe("foo/bar/baz.txt");
    });

    it("should normalize each segment", () => {
      expect(joinZipPath({}, "foo\\bar", "baz/qux")).toBe("foo/bar/baz/qux");
    });

    it("should skip empty segments", () => {
      expect(joinZipPath({}, "foo", "", "bar", "", "baz")).toBe("foo/bar/baz");
    });

    it("should handle single segment", () => {
      expect(joinZipPath({}, "foo")).toBe("foo");
    });

    it("should handle no segments", () => {
      expect(joinZipPath({})).toBe("");
    });

    it("should use provided mode for normalization", () => {
      expect(joinZipPath({ mode: "posix" }, "foo", "./bar", "baz/../qux")).toBe("foo/bar/qux");
    });

    it("should prepend slash when option is set", () => {
      expect(joinZipPath({ prependSlash: true }, "foo", "bar")).toBe("/foo/bar");
    });

    it("should strip leading slashes from segments", () => {
      expect(joinZipPath({}, "/foo", "/bar")).toBe("foo/bar");
    });

    it("should handle Windows paths", () => {
      expect(joinZipPath({ mode: "posix" }, "C:\\Users", "foo\\bar")).toBe("Users/foo/bar");
    });
  });
});
