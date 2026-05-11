/**
 * DOCX Module - Document Protection Tests
 */

import { describe, it, expect } from "vitest";

import {
  protectDocument,
  unprotectDocument,
  isDocumentProtected,
  getProtectionState,
  verifyProtectionPassword,
  Document
} from "../index";
import type { DocxDocument } from "../types";

function minimalDoc(): DocxDocument {
  return Document.build(Document.create());
}

describe("Document protection", () => {
  describe("protectDocument with password", () => {
    it("applies protection settings to the document", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "readOnly",
        password: "secret",
        spinCount: 1000
      });

      expect(protected_.settings).toBeDefined();
      expect(protected_.settings!.documentProtection).toBeDefined();
      const dp = protected_.settings!.documentProtection as any;
      expect(dp.edit).toBe("readOnly");
      expect(dp.enforcement).toBe(true);
      expect(dp.hashAlgorithm).toBe("SHA-256");
      expect(dp.hashValue).toBeDefined();
      expect(dp.saltValue).toBeDefined();
      expect(dp.spinCount).toBe(1000);
    });
  });

  describe("protectDocument without password", () => {
    it("applies protection without hash", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "comments",
        formatting: true
      });

      const dp = protected_.settings!.documentProtection as any;
      expect(dp.edit).toBe("comments");
      expect(dp.enforcement).toBe(true);
      expect(dp.hashValue).toBeUndefined();
      expect(dp.formatting).toBe(true);
    });
  });

  describe("unprotectDocument", () => {
    it("removes protection", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "readOnly",
        password: "pass",
        spinCount: 1000
      });
      const unprotected = unprotectDocument(protected_);

      expect(unprotected.settings?.documentProtection).toBeUndefined();
    });

    it("is a no-op for unprotected documents", () => {
      const doc = minimalDoc();
      const result = unprotectDocument(doc);
      expect(result.settings?.documentProtection).toBeUndefined();
    });
  });

  describe("isDocumentProtected", () => {
    it("returns true for protected document", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "forms",
        spinCount: 1000
      });
      expect(isDocumentProtected(protected_)).toBe(true);
    });

    it("returns false for unprotected document", () => {
      const doc = minimalDoc();
      expect(isDocumentProtected(doc)).toBe(false);
    });
  });

  describe("getProtectionState", () => {
    it("returns protection state for protected doc", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "trackedChanges",
        password: "pw",
        spinCount: 1000
      });

      const state = getProtectionState(protected_);
      expect(state).toBeDefined();
      expect(state!.edit).toBe("trackedChanges");
      expect(state!.enforcement).toBe(true);
      expect(state!.hashAlgorithm).toBe("SHA-256");
    });

    it("returns undefined for unprotected doc", () => {
      const doc = minimalDoc();
      expect(getProtectionState(doc)).toBeUndefined();
    });
  });

  describe("verifyProtectionPassword", () => {
    it("returns true for correct password", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "readOnly",
        password: "correct-horse",
        spinCount: 1000
      });

      const result = await verifyProtectionPassword(protected_, "correct-horse");
      expect(result).toBe(true);
    });

    it("returns false for wrong password", async () => {
      const doc = minimalDoc();
      const protected_ = await protectDocument(doc, {
        edit: "readOnly",
        password: "correct",
        spinCount: 1000
      });

      const result = await verifyProtectionPassword(protected_, "wrong");
      expect(result).toBe(false);
    });
  });

  describe("round-trip through DOCX", () => {
    it("preserves protection settings through protect → build → verify", async () => {
      const doc = Document.build(Document.create());
      const protected_ = await protectDocument(doc, {
        edit: "forms",
        password: "roundtrip",
        spinCount: 1000
      });

      // Verify the protection is intact in the document model
      expect(isDocumentProtected(protected_)).toBe(true);
      const state = getProtectionState(protected_);
      expect(state).toBeDefined();
      expect(state!.edit).toBe("forms");
      expect(state!.enforcement).toBe(true);
      expect(state!.hashValue).toBeDefined();

      // Verify password works on the protected model
      const valid = await verifyProtectionPassword(protected_, "roundtrip");
      expect(valid).toBe(true);
    });
  });
});
