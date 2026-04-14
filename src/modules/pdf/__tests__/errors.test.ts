import { PdfError, PdfRenderError, PdfFontError, PdfStructureError, isPdfError } from "@pdf/errors";
import { BaseError } from "@utils/errors";
/**
 * Tests for PDF error classes.
 */
import { describe, it, expect } from "vitest";

describe("PDF Errors", () => {
  describe("PdfError", () => {
    it("should extend BaseError", () => {
      const err = new PdfError("test");
      expect(err).toBeInstanceOf(BaseError);
      expect(err).toBeInstanceOf(Error);
    });

    it("should have correct name", () => {
      const err = new PdfError("test");
      expect(err.name).toBe("PdfError");
    });

    it("should preserve message", () => {
      const err = new PdfError("something went wrong");
      expect(err.message).toBe("something went wrong");
    });

    it("should support error cause chain", () => {
      const cause = new Error("root cause");
      const err = new PdfError("wrapper", { cause });
      expect(err.cause).toBe(cause);
    });
  });

  describe("PdfRenderError", () => {
    it("should extend PdfError", () => {
      const err = new PdfRenderError("render failed");
      expect(err).toBeInstanceOf(PdfError);
      expect(err.name).toBe("PdfRenderError");
    });
  });

  describe("PdfFontError", () => {
    it("should extend PdfError", () => {
      const err = new PdfFontError("font not found");
      expect(err).toBeInstanceOf(PdfError);
      expect(err.name).toBe("PdfFontError");
    });
  });

  describe("PdfStructureError", () => {
    it("should extend PdfError", () => {
      const err = new PdfStructureError("invalid structure");
      expect(err).toBeInstanceOf(PdfError);
      expect(err.name).toBe("PdfStructureError");
    });
  });

  describe("isPdfError", () => {
    it("should return true for PdfError", () => {
      expect(isPdfError(new PdfError("test"))).toBe(true);
    });

    it("should return true for PdfRenderError", () => {
      expect(isPdfError(new PdfRenderError("test"))).toBe(true);
    });

    it("should return true for PdfFontError", () => {
      expect(isPdfError(new PdfFontError("test"))).toBe(true);
    });

    it("should return true for PdfStructureError", () => {
      expect(isPdfError(new PdfStructureError("test"))).toBe(true);
    });

    it("should return false for regular Error", () => {
      expect(isPdfError(new Error("test"))).toBe(false);
    });

    it("should return false for null/undefined", () => {
      expect(isPdfError(null)).toBe(false);
      expect(isPdfError(undefined)).toBe(false);
    });
  });
});
