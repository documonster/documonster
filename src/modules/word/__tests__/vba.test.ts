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
  // Create fake binary data with "ThisDocument" string embedded
  const encoder = new TextEncoder();
  const prefix = new Uint8Array(50);
  const text = encoder.encode("ThisDocument");
  const suffix = new Uint8Array(50);
  const combined = new Uint8Array(prefix.length + text.length + suffix.length);
  combined.set(prefix, 0);
  combined.set(text, prefix.length);
  combined.set(suffix, prefix.length + text.length);
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

    it("returns true when vbaProject.bin is present", () => {
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
    it("adds VBA project to document", () => {
      const doc = makeDoc(undefined);
      const vbaData = createVbaData();
      const result = addVbaProject(doc, vbaData);

      expect(result.opaqueParts).toBeDefined();
      expect(result.opaqueParts!.some((p: any) => p.path === "word/vbaProject.bin")).toBe(true);
    });

    it("replaces existing VBA project", () => {
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: new Uint8Array(10) }]);
      const newVba = createVbaData();
      const result = addVbaProject(doc, newVba);

      const vbaParts = result.opaqueParts!.filter((p: any) => p.path === "word/vbaProject.bin");
      expect(vbaParts).toHaveLength(1);
      expect(vbaParts[0]!.data.length).toBe(newVba.length);
    });
  });

  describe("removeVbaProject", () => {
    it("removes VBA parts from document", () => {
      const doc = makeDoc([
        { path: "word/vbaProject.bin", data: createVbaData() },
        { path: "word/styles.xml", data: new Uint8Array(5) }
      ]);
      const result = removeVbaProject(doc);

      expect(result.opaqueParts!.some((p: any) => p.path.includes("vbaProject"))).toBe(false);
      expect(result.opaqueParts!.some((p: any) => p.path === "word/styles.xml")).toBe(true);
    });

    it("is a no-op for doc without VBA", () => {
      const doc = makeDoc(undefined);
      const result = removeVbaProject(doc);
      expect(result.opaqueParts).toBeUndefined();
    });
  });

  describe("listVbaParts", () => {
    it("returns empty for doc without VBA", () => {
      const doc = makeDoc(undefined);
      expect(listVbaParts(doc)).toHaveLength(0);
    });

    it("returns VBA-related parts", () => {
      const doc = makeDoc([
        { path: "word/vbaProject.bin", data: createVbaData() },
        { path: "word/vbaData.xml", data: new Uint8Array(5) },
        { path: "word/styles.xml", data: new Uint8Array(5) }
      ]);
      const parts = listVbaParts(doc);
      expect(parts.length).toBe(2);
    });
  });

  describe("getVbaProjectData", () => {
    it("returns undefined for doc without VBA", () => {
      const doc = makeDoc(undefined);
      expect(getVbaProjectData(doc)).toBeUndefined();
    });

    it("returns binary data for VBA project", () => {
      const vbaData = createVbaData();
      const doc = makeDoc([{ path: "word/vbaProject.bin", data: vbaData }]);
      const result = getVbaProjectData(doc);
      expect(result).toBeDefined();
      expect(result).toBe(vbaData);
    });
  });
});
