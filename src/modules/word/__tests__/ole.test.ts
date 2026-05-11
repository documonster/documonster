/**
 * DOCX Module - OLE Embedded Objects Tests
 */

import { describe, it, expect } from "vitest";

import { hasOleObjects, extractOleObjects, createOleEmbedding, getOleObjectData } from "../index";
import type { DocxDocument } from "../types";

function makeDoc(opaqueParts?: any[]): DocxDocument {
  return {
    body: [],
    sectionProperties: {},
    opaqueParts
  } as any;
}

describe("OLE embedded objects", () => {
  describe("hasOleObjects", () => {
    it("returns false for clean doc without opaque parts", () => {
      const doc = makeDoc(undefined);
      expect(hasOleObjects(doc)).toBe(false);
    });

    it("returns false for doc with non-OLE opaque parts", () => {
      const doc = makeDoc([{ path: "word/styles.xml", data: new Uint8Array(10) }]);
      expect(hasOleObjects(doc)).toBe(false);
    });

    it("returns true when embeddings are present", () => {
      const doc = makeDoc([{ path: "word/embeddings/oleObject1.bin", data: new Uint8Array(20) }]);
      expect(hasOleObjects(doc)).toBe(true);
    });
  });

  describe("extractOleObjects", () => {
    it("extracts metadata from OLE parts", () => {
      const doc = makeDoc([{ path: "word/embeddings/oleObject1.bin", data: new Uint8Array(50) }]);
      const result = extractOleObjects(doc);
      expect(result.objects.length).toBeGreaterThanOrEqual(1);
      expect(result.objects[0]!.objectType).toBe("embedded");
    });
  });

  describe("createOleEmbedding", () => {
    it("creates an opaque part for embedding", () => {
      const data = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
      const part = createOleEmbedding(data, "Excel.Sheet.12");
      expect(part.path).toContain("word/embeddings/");
      expect(part.data).toBe(data);
      expect(part.contentType).toContain("oleObject");
    });

    it("uses custom fileName", () => {
      const data = new Uint8Array(10);
      const part = createOleEmbedding(data, "Package", { fileName: "custom.bin" });
      expect(part.path).toBe("word/embeddings/custom.bin");
    });
  });

  describe("getOleObjectData", () => {
    it("returns undefined for missing OLE data", () => {
      const doc = makeDoc(undefined);
      expect(getOleObjectData(doc, "rId1")).toBeUndefined();
    });

    it("returns data for present OLE object", () => {
      const oleData = new Uint8Array([1, 2, 3, 4]);
      const doc = makeDoc([
        {
          path: "word/embeddings/oleObject1.bin",
          data: oleData,
          relationships: [{ id: "rId5", type: "ole", target: "embeddings/oleObject1.bin" }]
        }
      ]);
      const result = getOleObjectData(doc, "rId5");
      expect(result).toEqual(oleData);
    });
  });
});
