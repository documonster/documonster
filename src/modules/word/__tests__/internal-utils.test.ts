/**
 * DOCX Module - internal utility tests.
 *
 * Focus: `sanitizeUrl` — the security-critical helper that decides whether a
 * user-supplied URL is safe to emit verbatim into hyperlink relationships,
 * HTML output, Markdown link destinations, etc. A regression here could
 * reintroduce javascript:/data: URI injection, so the rules deserve direct
 * coverage rather than relying on indirect tests in conversion modules.
 */

import { describe, it, expect } from "vitest";

import { sanitizeUrl } from "../core/internal-utils";

describe("sanitizeUrl", () => {
  describe("nullish & empty", () => {
    it("returns undefined for null", () => {
      expect(sanitizeUrl(null)).toBeUndefined();
    });
    it("returns undefined for undefined", () => {
      expect(sanitizeUrl(undefined)).toBeUndefined();
    });
    it("returns undefined for empty string", () => {
      expect(sanitizeUrl("")).toBeUndefined();
    });
    it("returns undefined for whitespace-only string", () => {
      expect(sanitizeUrl("   \t\n  ")).toBeUndefined();
    });
  });

  describe("safe schemes", () => {
    it("accepts http", () => {
      expect(sanitizeUrl("http://example.com/x")).toBe("http://example.com/x");
    });
    it("accepts https", () => {
      expect(sanitizeUrl("https://example.com/x")).toBe("https://example.com/x");
    });
    it("accepts mailto", () => {
      expect(sanitizeUrl("mailto:a@b.c")).toBe("mailto:a@b.c");
    });
    it("accepts tel", () => {
      expect(sanitizeUrl("tel:+1-555-0100")).toBe("tel:+1-555-0100");
    });
    it("accepts ftp / ftps", () => {
      expect(sanitizeUrl("ftp://host/path")).toBe("ftp://host/path");
      expect(sanitizeUrl("ftps://host/path")).toBe("ftps://host/path");
    });
    it("accepts sms", () => {
      expect(sanitizeUrl("sms:+15550100")).toBe("sms:+15550100");
    });
    it("accepts uppercase scheme (case-insensitive)", () => {
      expect(sanitizeUrl("HTTPS://example.com")).toBe("HTTPS://example.com");
      expect(sanitizeUrl("MailTo:a@b.c")).toBe("MailTo:a@b.c");
    });
  });

  describe("dangerous schemes", () => {
    it("rejects javascript:", () => {
      expect(sanitizeUrl("javascript:alert(1)")).toBeUndefined();
    });
    it("rejects vbscript:", () => {
      expect(sanitizeUrl("vbscript:msgbox(1)")).toBeUndefined();
    });
    it("rejects data: URIs", () => {
      expect(
        sanitizeUrl("data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==")
      ).toBeUndefined();
    });
    it("rejects file:", () => {
      expect(sanitizeUrl("file:///etc/passwd")).toBeUndefined();
    });
    it("rejects unknown scheme", () => {
      expect(sanitizeUrl("chrome-extension://abc/page")).toBeUndefined();
    });
  });

  describe("scheme-detection bypass attempts", () => {
    it("rejects javascript: with leading whitespace", () => {
      expect(sanitizeUrl("   javascript:alert(1)")).toBeUndefined();
    });
    it("rejects javascript: with embedded tab (browsers strip control chars before parsing)", () => {
      expect(sanitizeUrl("java\tscript:alert(1)")).toBeUndefined();
    });
    it("rejects javascript: with embedded newline", () => {
      expect(sanitizeUrl("java\nscript:alert(1)")).toBeUndefined();
    });
    it("rejects javascript: with embedded NUL", () => {
      expect(sanitizeUrl("java\u0000script:alert(1)")).toBeUndefined();
    });
    it("rejects javascript: with mixed control characters", () => {
      expect(sanitizeUrl(" \u0001JAVA\u0002SCRIPT\u0003:alert(1)")).toBeUndefined();
    });
  });

  describe("non-absolute URLs", () => {
    it("keeps fragment-only URLs", () => {
      expect(sanitizeUrl("#anchor")).toBe("#anchor");
    });
    it("keeps protocol-relative URLs", () => {
      expect(sanitizeUrl("//cdn.example.com/x.js")).toBe("//cdn.example.com/x.js");
    });
    it("keeps relative paths", () => {
      expect(sanitizeUrl("./local/file")).toBe("./local/file");
      expect(sanitizeUrl("../up/file")).toBe("../up/file");
      expect(sanitizeUrl("relative/path")).toBe("relative/path");
    });
    it("keeps query-only URLs", () => {
      expect(sanitizeUrl("?q=1&r=2")).toBe("?q=1&r=2");
    });
    it("trims surrounding whitespace from relative URLs", () => {
      expect(sanitizeUrl("  ./x  ")).toBe("./x");
    });
  });

  describe("trimming & normalization", () => {
    it("trims leading/trailing whitespace from accepted URLs", () => {
      expect(sanitizeUrl("  https://example.com  ")).toBe("https://example.com");
    });
    it("strips surrounding control characters", () => {
      expect(sanitizeUrl("\u0000https://example.com\u0001")).toBe("https://example.com");
    });
  });
});
