/**
 * DOCX Module - Document Protection Tests
 */

import { describe, it, expect } from "vitest";

import { Document, Security } from "../index";
import type { DocxDocument } from "../types";

function minimalDoc(): DocxDocument {
  return Document.build(Document.create());
}

describe("Document protection", () => {
  describe("protectDocument with password", () => {
    it("applies protection settings to the document", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
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

    it("computes the ISO/IEC 29500 password hash (iterator appended, not prepended)", async () => {
      // Regression guard: Word's documentProtection hash is
      //   Hi = Hash(Hi-1 + LE_uint32(i))
      // i.e. the iterator is appended AFTER the previous hash. If the order is
      // flipped, Word cannot reproduce the hash and treats the document as
      // unprotected (offering "Start Enforcing Protection" instead of
      // prompting for the password). Reproduce the spec hash here with the
      // same salt protectDocument used and compare.
      const { hashAsync } = await import("@utils/crypto");
      const { stringToUtf16LE, base64ToBytes, bytesToBase64 } =
        await import("../core/internal-utils");

      const password = "swordfish";
      const algo = "SHA-512";
      const spin = 1000;

      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
        edit: "readOnly",
        password,
        hashAlgorithm: algo,
        spinCount: spin
      });
      const dp = protected_.settings!.documentProtection as any;

      // Independent reference implementation of the ISO/IEC 29500 hash.
      const saltBytes = base64ToBytes(dp.saltValue);
      const pw = stringToUtf16LE(password);
      const init = new Uint8Array(saltBytes.length + pw.length);
      init.set(saltBytes, 0);
      init.set(pw, saltBytes.length);
      let h = await hashAsync(algo, init);
      for (let i = 0; i < spin; i++) {
        const it = new Uint8Array(4);
        it[0] = i & 0xff;
        it[1] = (i >> 8) & 0xff;
        it[2] = (i >> 16) & 0xff;
        it[3] = (i >> 24) & 0xff;
        const c = new Uint8Array(h.length + 4);
        c.set(h, 0);
        c.set(it, h.length);
        h = await hashAsync(algo, c);
      }
      expect(dp.hashValue).toBe(bytesToBase64(h));

      // And the verifier accepts the correct password / rejects a wrong one.
      expect(await Security.verifyPassword(protected_, "swordfish")).toBe(true);
      expect(await Security.verifyPassword(protected_, "wrong")).toBe(false);
    });
  });

  describe("protectDocument without password", () => {
    it("applies protection without hash", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
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
      const protected_ = await Security.protect(doc, {
        edit: "readOnly",
        password: "pass",
        spinCount: 1000
      });
      const unprotected = Security.unprotect(protected_);

      expect(unprotected.settings?.documentProtection).toBeUndefined();
    });

    it("is a no-op for unprotected documents", () => {
      const doc = minimalDoc();
      const result = Security.unprotect(doc);
      expect(result.settings?.documentProtection).toBeUndefined();
    });
  });

  describe("isDocumentProtected", () => {
    it("returns true for protected document", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
        edit: "forms",
        spinCount: 1000
      });
      expect(Security.isProtected(protected_)).toBe(true);
    });

    it("returns false for unprotected document", () => {
      const doc = minimalDoc();
      expect(Security.isProtected(doc)).toBe(false);
    });
  });

  describe("getProtectionState", () => {
    it("returns protection state for protected doc", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
        edit: "trackedChanges",
        password: "pw",
        spinCount: 1000
      });

      const state = Security.getState(protected_);
      expect(state).toBeDefined();
      expect(state!.edit).toBe("trackedChanges");
      expect(state!.enforcement).toBe(true);
      expect(state!.hashAlgorithm).toBe("SHA-256");
    });

    it("returns undefined for unprotected doc", () => {
      const doc = minimalDoc();
      expect(Security.getState(doc)).toBeUndefined();
    });
  });

  describe("verifyProtectionPassword", () => {
    it("returns true for correct password", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
        edit: "readOnly",
        password: "correct-horse",
        spinCount: 1000
      });

      const result = await Security.verifyPassword(protected_, "correct-horse");
      expect(result).toBe(true);
    });

    it("returns false for wrong password", async () => {
      const doc = minimalDoc();
      const protected_ = await Security.protect(doc, {
        edit: "readOnly",
        password: "correct",
        spinCount: 1000
      });

      const result = await Security.verifyPassword(protected_, "wrong");
      expect(result).toBe(false);
    });
  });

  describe("round-trip through DOCX", () => {
    it("preserves protection settings through protect → build → verify", async () => {
      const doc = Document.build(Document.create());
      const protected_ = await Security.protect(doc, {
        edit: "forms",
        password: "roundtrip",
        spinCount: 1000
      });

      // Verify the protection is intact in the document model
      expect(Security.isProtected(protected_)).toBe(true);
      const state = Security.getState(protected_);
      expect(state).toBeDefined();
      expect(state!.edit).toBe("forms");
      expect(state!.enforcement).toBe(true);
      expect(state!.hashValue).toBeDefined();

      // Verify password works on the protected model
      const valid = await Security.verifyPassword(protected_, "roundtrip");
      expect(valid).toBe(true);
    });
  });
});
