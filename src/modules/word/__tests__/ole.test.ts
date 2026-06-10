/**
 * DOCX Module - OLE Embedded Objects Tests
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  hasOleObjects,
  extractOleObjects,
  createOleEmbedding,
  getOleObjectData,
  addOleObject,
  toBuffer,
  readDocx
} from "../index";
import type { DocxDocument } from "../types";

/** A minimal real OLE2 compound-document header so progId detection works. */
const OLE2_HEADER = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

/** 1×1 PNG used as an OLE preview image in round-trip tests. */
const PNG_1X1 = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

function makeBaseDoc(): DocxDocument {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "Body with an OLE object.");
  return Document.build(d);
}

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
      const result = createOleEmbedding(data, "Excel.Sheet.12");
      expect(result.olePart.path).toContain("word/embeddings/");
      expect(result.olePart.data).toBe(data);
      expect(result.olePart.contentType).toContain("oleObject");
      expect(result.oleRId).toBeTruthy();
      expect(result.previewPart).toBeUndefined();
    });

    it("uses custom fileName", () => {
      const data = new Uint8Array(10);
      const result = createOleEmbedding(data, "Package", { fileName: "custom.bin" });
      expect(result.olePart.path).toBe("word/embeddings/custom.bin");
    });

    it("allocates unique default file names across calls", () => {
      const a = createOleEmbedding(new Uint8Array(1), "X");
      const b = createOleEmbedding(new Uint8Array(1), "X");
      expect(a.olePart.path).not.toBe(b.olePart.path);
    });

    it("emits a preview media part when previewImage is provided", () => {
      const result = createOleEmbedding(new Uint8Array(1), "Excel.Sheet.12", {
        previewImage: new Uint8Array([1, 2, 3]),
        previewContentType: "image/png"
      });
      expect(result.previewPart).toBeDefined();
      expect(result.previewPart!.path).toMatch(/^word\/media\/.+\.png$/);
      expect(result.previewRId).toBeTruthy();
    });

    it("throws when previewImage is given without previewContentType", () => {
      expect(() =>
        createOleEmbedding(new Uint8Array(1), "X", {
          previewImage: new Uint8Array(1)
        })
      ).toThrow(/previewContentType/);
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

  describe("addOleObject + round-trip", () => {
    it("wires the OLE object, document relationship and body reference", async () => {
      const ole = createOleEmbedding(OLE2_HEADER, "Excel.Sheet.12");
      const doc = addOleObject(makeBaseDoc(), ole);

      // Model side: oleObjects field carries the exact rId + progId.
      expect(doc.oleObjects).toHaveLength(1);
      expect(doc.oleObjects![0]!.rId).toBe(ole.oleRId);
      expect(doc.oleObjects![0]!.progId).toBe("Excel.Sheet.12");
      // Body side: a <w:object>/<o:OLEObject> referencing the same rId.
      expect(hasOleObjects(doc)).toBe(true);
      expect(getOleObjectData(doc, ole.oleRId)).toEqual(OLE2_HEADER);

      // Round-trip through the package.
      const reread = await readDocx(await toBuffer(doc));
      expect(hasOleObjects(reread)).toBe(true);
      expect(getOleObjectData(reread, ole.oleRId)).toEqual(OLE2_HEADER);

      const extraction = extractOleObjects(reread);
      expect(extraction.objects).toHaveLength(1);
      expect(extraction.objects[0]!.rId).toBe(ole.oleRId);
      // progId survives because it is encoded in the body <o:OLEObject ProgID>.
      expect(extraction.objects[0]!.progId).toBe("Excel.Sheet.12");
    });

    it("wires a preview image with its own relationship", async () => {
      const ole = createOleEmbedding(OLE2_HEADER, "Excel.Sheet.12", {
        previewImage: PNG_1X1,
        previewContentType: "image/png"
      });
      const doc = addOleObject(makeBaseDoc(), ole);
      expect(doc.oleObjects![0]!.previewRId).toBe(ole.previewRId);

      const buf = await toBuffer(doc);
      const reread = await readDocx(buf);
      expect(getOleObjectData(reread, ole.oleRId)).toEqual(OLE2_HEADER);
      // The OLE binary's rId never collides with the preview's rId.
      expect(extractOleObjects(reread).objects[0]!.rId).toBe(ole.oleRId);
    });

    it("strips OLE binaries when preserveOleObjects is false", async () => {
      const ole = createOleEmbedding(OLE2_HEADER, "Excel.Sheet.12");
      const doc = addOleObject(makeBaseDoc(), ole);
      const buf = await toBuffer(doc);
      const reread = await readDocx(buf, { securityPolicy: { preserveOleObjects: false } });
      expect(hasOleObjects(reread)).toBe(false);
      expect(getOleObjectData(reread, ole.oleRId)).toBeUndefined();
    });
  });
});
