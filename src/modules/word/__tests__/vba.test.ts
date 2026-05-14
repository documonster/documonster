/**
 * DOCX Module - VBA Project Tests
 */

import { describe, it, expect } from "vitest";

import {
  hasVbaProject,
  getVbaProjectInfo,
  addVbaProject,
  removeVbaProject,
  listVbaParts,
  getVbaProjectData
} from "../index";
import type { DocxDocument } from "../types";

function makeDoc(opaqueParts?: any[]): DocxDocument {
  return {
    body: [],
    sectionProperties: {},
    opaqueParts
  } as any;
}

function createVbaData(): Uint8Array {
  // OLE2 compound document signature, then "ThisDocument" string embedded
  // for the heuristic module-name scanner.
  const encoder = new TextEncoder();
  const ole2 = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  const prefix = new Uint8Array(50);
  const text = encoder.encode("ThisDocument");
  const suffix = new Uint8Array(50);
  const combined = new Uint8Array(ole2.length + prefix.length + text.length + suffix.length);
  combined.set(ole2, 0);
  combined.set(prefix, ole2.length);
  combined.set(text, ole2.length + prefix.length);
  combined.set(suffix, ole2.length + prefix.length + text.length);
  return combined;
}

describe("VBA Project", () => {
  describe("hasVbaProject", () => {
    it("returns false for document without VBA", () => {
      const doc = makeDoc(undefined);
      expect(hasVbaProject(doc)).toBe(false);
    });

    it("returns false for doc with non-VBA opaque parts", () => {
      const doc = makeDoc([{ path: "word/styles.xml", data: new Uint8Array(10) }]);
      expect(hasVbaProject(doc)).toBe(false);
    });

    it("returns true when vbaProject is present on the canonical field", () => {
      const doc = { ...makeDoc(undefined), vbaProject: createVbaData() } as DocxDocument;
      expect(hasVbaProject(doc)).toBe(true);
    });

    it("returns true for legacy opaqueParts placement", () => {
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: createVbaData() }]);
      expect(hasVbaProject(doc)).toBe(true);
    });
  });

  describe("getVbaProjectInfo", () => {
    it("returns hasVba=false for clean doc", () => {
      const doc = makeDoc(undefined);
      const info = getVbaProjectInfo(doc);
      expect(info.hasVba).toBe(false);
    });

    it("extracts info with ThisDocument module name", () => {
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: createVbaData() }]);
      const info = getVbaProjectInfo(doc);
      expect(info.hasVba).toBe(true);
      expect(info.projectPath).toBe("word/vbaProject.bin");
      expect(info.moduleNames).toContain("ThisDocument");
      expect(info.sizeBytes).toBeGreaterThan(0);
    });
  });

  describe("addVbaProject", () => {
    it("stores binary on the canonical doc.vbaProject field", () => {
      const doc = makeDoc(undefined);
      const vbaData = createVbaData();
      const result = addVbaProject(doc, vbaData);
      expect(result.vbaProject).toBeDefined();
      expect(result.vbaProject).toBe(vbaData);
    });

    it("rejects data without an OLE2 compound document header", () => {
      const doc = makeDoc(undefined);
      // No OLE2 prefix → must throw.
      expect(() => addVbaProject(doc, new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]))).toThrow(
        /OLE2/
      );
    });

    it("removes legacy opaqueParts copies of vbaProject.bin", () => {
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: new Uint8Array(10) }]);
      const newVba = createVbaData();
      const result = addVbaProject(doc, newVba);
      // Canonical field has the new data.
      expect(result.vbaProject).toBe(newVba);
      // No leftover opaque copy that the packager would now reject.
      const stale = result.opaqueParts?.some(p => p.path === "word/vbaProject.bin");
      expect(stale ?? false).toBe(false);
    });

    it("promotes a plain document to macroEnabledDocument so the .docm content type matches", () => {
      // Word rejects packages that carry a vbaProject.bin but declare
      // their main part as the plain wordprocessingml document type.
      const doc = makeDoc(undefined);
      expect(doc.docType).toBeUndefined();
      const result = addVbaProject(doc, createVbaData());
      expect(result.docType).toBe("macroEnabledDocument");
    });

    it("promotes a template to macroEnabledTemplate", () => {
      const doc: DocxDocument = { ...makeDoc(undefined), docType: "template" };
      const result = addVbaProject(doc, createVbaData());
      expect(result.docType).toBe("macroEnabledTemplate");
    });

    it("preserves an already-macroEnabled docType unchanged", () => {
      const doc: DocxDocument = {
        ...makeDoc(undefined),
        docType: "macroEnabledDocument"
      };
      const result = addVbaProject(doc, createVbaData());
      expect(result.docType).toBe("macroEnabledDocument");
    });
  });

  describe("removeVbaProject", () => {
    it("removes VBA parts from document", () => {
      const doc = {
        ...makeDoc([
          { path: "word/vbaProject.bin", data: createVbaData() },
          { path: "word/styles.xml", data: new Uint8Array(5) }
        ]),
        vbaProject: createVbaData()
      } as DocxDocument;
      const result = removeVbaProject(doc);

      expect(result.vbaProject).toBeUndefined();
      expect(result.opaqueParts!.some((p: any) => p.path.includes("vbaProject"))).toBe(false);
      expect(result.opaqueParts!.some((p: any) => p.path === "word/styles.xml")).toBe(true);
    });

    it("is a no-op for doc without VBA", () => {
      const doc = makeDoc(undefined);
      const result = removeVbaProject(doc);
      expect(result.vbaProject).toBeUndefined();
      expect(result.opaqueParts).toBeUndefined();
    });

    it("demotes macroEnabledDocument back to document on removal", () => {
      const doc: DocxDocument = {
        ...makeDoc(undefined),
        docType: "macroEnabledDocument",
        vbaProject: createVbaData()
      };
      const result = removeVbaProject(doc);
      expect(result.docType).toBe("document");
      expect(result.vbaProject).toBeUndefined();
    });

    it("demotes macroEnabledTemplate back to template on removal", () => {
      const doc: DocxDocument = {
        ...makeDoc(undefined),
        docType: "macroEnabledTemplate",
        vbaProject: createVbaData()
      };
      const result = removeVbaProject(doc);
      expect(result.docType).toBe("template");
    });
  });

  describe("listVbaParts", () => {
    it("returns empty for doc without VBA", () => {
      const doc = makeDoc(undefined);
      expect(listVbaParts(doc)).toHaveLength(0);
    });

    it("returns the canonical vbaProject field plus auxiliary parts", () => {
      const vbaData = createVbaData();
      const doc = {
        ...makeDoc([
          { path: "word/vbaData.xml", data: new Uint8Array(5) },
          { path: "word/styles.xml", data: new Uint8Array(5) }
        ]),
        vbaProject: vbaData
      } as DocxDocument;
      const parts = listVbaParts(doc);
      expect(parts.length).toBe(2);
      const paths = parts.map(p => p.path).sort();
      expect(paths).toEqual(["word/vbaData.xml", "word/vbaProject.bin"]);
    });
  });

  describe("getVbaProjectData", () => {
    it("returns undefined for doc without VBA", () => {
      const doc = makeDoc(undefined);
      expect(getVbaProjectData(doc)).toBeUndefined();
    });

    it("returns binary data for VBA project on the canonical field", () => {
      const vbaData = createVbaData();
      const doc = { ...makeDoc(undefined), vbaProject: vbaData } as DocxDocument;
      const result = getVbaProjectData(doc);
      expect(result).toBe(vbaData);
    });

    it("falls back to opaqueParts for hand-built models", () => {
      const vbaData = createVbaData();
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: vbaData }]);
      const result = getVbaProjectData(doc);
      expect(result).toBe(vbaData);
    });
  });
});
