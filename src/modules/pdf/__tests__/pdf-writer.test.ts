/**
 * Tests for the PDF Writer (file assembly).
 */
import { describe, it, expect } from "vitest";
import { PdfWriter } from "@pdf/core/pdf-writer";
import { PdfDict, pdfRef } from "@pdf/core/pdf-object";
import { PdfContentStream } from "@pdf/core/pdf-stream";
import { PdfStructureError } from "@pdf/errors";
import { initEncryption } from "@pdf/core/encryption";
import { alphaGsName } from "@pdf/render/page-renderer";

describe("PdfWriter", () => {
  describe("Object Allocation", () => {
    it("should allocate sequential object numbers", () => {
      const writer = new PdfWriter();
      expect(writer.allocObject()).toBe(1);
      expect(writer.allocObject()).toBe(2);
      expect(writer.allocObject()).toBe(3);
    });
  });

  describe("build()", () => {
    it("should throw without catalog", () => {
      const writer = new PdfWriter();
      expect(() => writer.build()).toThrow(PdfStructureError);
    });

    it("should produce valid PDF structure", () => {
      const writer = new PdfWriter();

      // Create a minimal valid PDF
      const contentObjNum = writer.allocObject();
      const stream = new PdfContentStream();
      stream.beginText().setFont("F1", 12).showText("Hello").endText();
      const contentDict = new PdfDict();
      writer.addStreamObject(contentObjNum, contentDict, stream);

      // Font resource
      const fontObjNum = writer.allocObject();
      const fontDict = new PdfDict()
        .set("Type", "/Font")
        .set("Subtype", "/Type1")
        .set("BaseFont", "/Helvetica");
      writer.addObject(fontObjNum, fontDict);

      // Resources
      const resourcesObjNum = writer.allocObject();
      const resourcesDict = new PdfDict().set("Font", `<< /F1 ${pdfRef(fontObjNum)} >>`);
      writer.addObject(resourcesObjNum, resourcesDict);

      // Page tree
      const pagesObjNum = writer.allocObject();

      // Page
      const pageObjNum = writer.addPage({
        parentRef: pagesObjNum,
        width: 595.28,
        height: 841.89,
        contentsRef: contentObjNum,
        resourcesRef: resourcesObjNum
      });

      // Fill in page tree
      const pagesDict = new PdfDict()
        .set("Type", "/Pages")
        .set("Kids", `[${pdfRef(pageObjNum)}]`)
        .set("Count", "1");
      writer.addObject(pagesObjNum, pagesDict);

      // Catalog
      writer.addCatalog(pagesObjNum);

      // Build
      const pdf = writer.build();
      expect(pdf).toBeInstanceOf(Uint8Array);
      expect(pdf.length).toBeGreaterThan(0);

      // Verify PDF structure
      const text = new TextDecoder().decode(pdf);
      expect(text).toContain("%PDF-2.0");
      expect(text).toContain("%%EOF");
      expect(text).toContain("xref");
      expect(text).toContain("trailer");
      expect(text).toContain("/Catalog");
      expect(text).toContain("/Pages");
      expect(text).toContain("/Page");
      expect(text).toContain("stream");
      expect(text).toContain("endstream");
      expect(text).toContain("/Helvetica");
    });

    it("should include info dictionary", () => {
      const writer = new PdfWriter();

      // Minimal structure
      const contentObjNum = writer.allocObject();
      writer.addStreamObject(contentObjNum, new PdfDict(), new PdfContentStream());

      const resourcesObjNum = writer.allocObject();
      writer.addObject(resourcesObjNum, new PdfDict());

      const pagesObjNum = writer.allocObject();
      const pageObjNum = writer.addPage({
        parentRef: pagesObjNum,
        width: 595.28,
        height: 841.89,
        contentsRef: contentObjNum,
        resourcesRef: resourcesObjNum
      });

      writer.addObject(
        pagesObjNum,
        new PdfDict()
          .set("Type", "/Pages")
          .set("Kids", `[${pdfRef(pageObjNum)}]`)
          .set("Count", "1")
      );

      writer.addCatalog(pagesObjNum);
      writer.addInfoDict({
        title: "Test Doc",
        author: "Test Author",
        creator: "excelts"
      });

      const pdf = writer.build();
      const text = new TextDecoder().decode(pdf);
      expect(text).toContain("/Title (Test Doc)");
      expect(text).toContain("/Author (Test Author)");
      expect(text).toContain("/Producer (excelts)");
      expect(text).toContain("/Info");
    });
  });

  describe("xref free list conformance", () => {
    it("should use generation 65535 for free entries", () => {
      const writer = new PdfWriter();

      // Allocate objects 1, 2, 3 but only add 1 and 3 — object 2 is a gap
      const obj1 = writer.allocObject(); // 1
      const _obj2 = writer.allocObject(); // 2 — gap
      const obj3 = writer.allocObject(); // 3

      writer.addObject(obj1, new PdfDict().set("Type", "/Font"));
      writer.addObject(obj3, new PdfDict().set("Type", "/Font"));

      const pagesObjNum = writer.allocObject();
      const pageObjNum = writer.addPage({
        parentRef: pagesObjNum,
        width: 595,
        height: 842,
        contentsRef: obj1,
        resourcesRef: obj3
      });
      writer.addObject(
        pagesObjNum,
        new PdfDict()
          .set("Type", "/Pages")
          .set("Kids", `[${pdfRef(pageObjNum)}]`)
          .set("Count", "1")
      );
      writer.addCatalog(pagesObjNum);

      const pdf = writer.build();
      const text = new TextDecoder().decode(pdf);

      // Object 0 free entry should have generation 65535
      expect(text).toMatch(/0000000\d+ 65535 f /);
      // Gap entries (object 2) should also have generation 65535
      const xrefSection = text.slice(text.indexOf("xref"), text.indexOf("trailer"));
      expect(xrefSection).not.toContain("00000 f ");
      // All free entries should have 65535
      const freeEntries = xrefSection.match(/\d{10} \d{5} f /g) ?? [];
      for (const entry of freeEntries) {
        expect(entry).toContain("65535 f ");
      }
    });
  });

  describe("encryption stream encryption", () => {
    it("should encrypt stream data at build time", () => {
      const writer = new PdfWriter();

      const contentObjNum = writer.allocObject();
      const stream = new PdfContentStream();
      stream.beginText().setFont("F1", 12).showText("Hello").endText();
      writer.addStreamObject(contentObjNum, new PdfDict(), stream);

      const resourcesObjNum = writer.allocObject();
      writer.addObject(resourcesObjNum, new PdfDict());

      const pagesObjNum = writer.allocObject();
      const pageObjNum = writer.addPage({
        parentRef: pagesObjNum,
        width: 595,
        height: 842,
        contentsRef: contentObjNum,
        resourcesRef: resourcesObjNum
      });
      writer.addObject(
        pagesObjNum,
        new PdfDict()
          .set("Type", "/Pages")
          .set("Kids", `[${pdfRef(pageObjNum)}]`)
          .set("Count", "1")
      );
      writer.addCatalog(pagesObjNum);

      // Enable encryption AFTER all objects are added (this is the real-world order)
      const encState = initEncryption({ ownerPassword: "test" });
      writer.setEncryption(encState);

      const pdf = writer.build();
      const text = new TextDecoder().decode(pdf);

      // Should have encryption dict
      expect(text).toContain("/Filter /Standard");
      expect(text).toContain("/Encrypt");

      // The stream content should be encrypted — the original "Hello" text
      // (even compressed) should not appear in the same form as without encryption
      const pdfPlain = (() => {
        const w2 = new PdfWriter();
        const co = w2.allocObject();
        const s2 = new PdfContentStream();
        s2.beginText().setFont("F1", 12).showText("Hello").endText();
        w2.addStreamObject(co, new PdfDict(), s2);
        const ro = w2.allocObject();
        w2.addObject(ro, new PdfDict());
        const po = w2.allocObject();
        const pg = w2.addPage({
          parentRef: po,
          width: 595,
          height: 842,
          contentsRef: co,
          resourcesRef: ro
        });
        w2.addObject(
          po,
          new PdfDict()
            .set("Type", "/Pages")
            .set("Kids", `[${pdfRef(pg)}]`)
            .set("Count", "1")
        );
        w2.addCatalog(po);
        return w2.build();
      })();

      // The encrypted PDF should differ from the plain one
      expect(pdf.length).not.toBe(pdfPlain.length);
    });

    it("should preserve literal backslash sequences when encrypting strings", () => {
      const writer = new PdfWriter();
      const objNum = writer.allocObject();
      writer.addObject(objNum, new PdfDict().set("Title", "(hello\\\\nworld)"));

      const pagesObjNum = writer.allocObject();
      const pageObjNum = writer.addPage({
        parentRef: pagesObjNum,
        width: 595,
        height: 842,
        contentsRef: objNum,
        resourcesRef: objNum
      });
      writer.addObject(
        pagesObjNum,
        new PdfDict()
          .set("Type", "/Pages")
          .set("Kids", `[${pdfRef(pageObjNum)}]`)
          .set("Count", "1")
      );
      writer.addCatalog(pagesObjNum);

      writer.setEncryption(initEncryption({ ownerPassword: "test" }));
      const pdf = writer.build();
      const text = new TextDecoder().decode(pdf);

      expect(text).toContain("/Encrypt");
      expect(text).not.toContain("(hello\\\\nworld)");
      expect(text).toMatch(/\/Title <[0-9a-f]+>/i);
    });
  });
});

describe("alphaGsName", () => {
  it("should produce different names for close alpha values", () => {
    expect(alphaGsName(0.504)).not.toBe(alphaGsName(0.506));
  });

  it("should produce deterministic names", () => {
    expect(alphaGsName(0.5)).toBe(alphaGsName(0.5));
  });

  it("should use 4-digit precision", () => {
    expect(alphaGsName(0.5)).toBe("GS5000");
    expect(alphaGsName(0.123)).toBe("GS1230");
  });
});
