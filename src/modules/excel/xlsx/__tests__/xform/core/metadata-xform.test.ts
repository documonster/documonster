import { MetadataXform } from "@excel/xlsx/xform/core/metadata-xform";
import { describe, expect, it } from "vitest";

function xmlToStream(xml: string) {
  const encoder = new TextEncoder();
  const data = encoder.encode(xml);
  return {
    [Symbol.asyncIterator]() {
      let done = false;
      return {
        next() {
          if (done) {
            return Promise.resolve({ done: true as const, value: undefined });
          }
          done = true;
          return Promise.resolve({ done: false as const, value: data });
        }
      };
    }
  } as any;
}

describe("MetadataXform", () => {
  describe("render", () => {
    it("should generate valid XLDAPR metadata XML when dynamicArrayCount > 0", () => {
      const xform = new MetadataXform();
      const xml = xform.toXml({ dynamicArrayCount: 3 });

      expect(xml).toContain("<metadata");
      expect(xml).toContain(
        'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"'
      );
      expect(xml).toContain('<metadataType name="XLDAPR"');
      expect(xml).toContain('cellMeta="1"');
      expect(xml).toContain('<futureMetadata name="XLDAPR" count="1"');
      expect(xml).toContain("<xda:dynamicArrayProperties");
      expect(xml).toContain('fDynamic="1"');
      expect(xml).toContain('fCollapsed="0"');
      expect(xml).toContain('<cellMetadata count="1"');
      expect(xml).toContain('<rc t="1" v="0"');
    });

    it("should produce empty output when dynamicArrayCount is 0", () => {
      const xform = new MetadataXform();
      const xml = xform.toXml({ dynamicArrayCount: 0 });
      expect(xml).toBe("");
    });

    it("should produce empty output when model is null", () => {
      const xform = new MetadataXform();
      const xml = xform.toXml(null as any);
      expect(xml).toBe("");
    });
  });

  describe("parse", () => {
    it("should detect XLDAPR dynamic array metadata", async () => {
      const xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        '  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">',
        '  <metadataTypes count="1">',
        '    <metadataType name="XLDAPR" minSupportedVersion="120000"',
        '      copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1"',
        '      rowColShift="1" clearFormats="1" clearComments="1" assign="1"',
        '      coerce="1" adjust="1" cellMeta="1"/>',
        "  </metadataTypes>",
        '  <futureMetadata name="XLDAPR" count="1">',
        "    <bk>",
        "      <extLst>",
        '        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">',
        '          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
        "        </ext>",
        "      </extLst>",
        "    </bk>",
        "  </futureMetadata>",
        '  <cellMetadata count="1">',
        "    <bk>",
        '      <rc t="1" v="0"/>',
        "    </bk>",
        "  </cellMetadata>",
        "</metadata>"
      ].join("\n");

      const xform = new MetadataXform();
      const result = await xform.parseStream(xmlToStream(xml));
      expect(result).toBeDefined();
      expect(result.hasDynamicArrays).toBe(true);
    });

    it("should return hasDynamicArrays=false for metadata without XLDAPR", async () => {
      const xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '  <metadataTypes count="0"/>',
        '  <cellMetadata count="0"/>',
        "</metadata>"
      ].join("\n");

      const xform = new MetadataXform();
      const result = await xform.parseStream(xmlToStream(xml));
      expect(result.hasDynamicArrays).toBe(false);
      expect(result.dynamicArrayCmIndices.size).toBe(0);
    });

    it("should build precise dynamicArrayCmIndices for XLDAPR cellMetadata", async () => {
      const xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        '  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">',
        '  <metadataTypes count="1">',
        '    <metadataType name="XLDAPR" minSupportedVersion="120000" cellMeta="1"/>',
        "  </metadataTypes>",
        '  <futureMetadata name="XLDAPR" count="1">',
        '    <bk><extLst><ext><xda:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk>',
        "  </futureMetadata>",
        '  <cellMetadata count="1">',
        "    <bk>",
        '      <rc t="1" v="0"/>',
        "    </bk>",
        "  </cellMetadata>",
        "</metadata>"
      ].join("\n");

      const xform = new MetadataXform();
      const result = await xform.parseStream(xmlToStream(xml));
      expect(result.hasDynamicArrays).toBe(true);
      // cm=1 (1-indexed) should map to XLDAPR
      expect(result.dynamicArrayCmIndices.has(1)).toBe(true);
      expect(result.dynamicArrayCmIndices.size).toBe(1);
    });

    it("should not include cm indices for non-XLDAPR metadata types", async () => {
      // Two metadata types: XLRICHVALUE (index 1) and XLDAPR (index 2)
      // cellMetadata has two bk entries:
      //   bk[0] → rc t="1" (XLRICHVALUE) → cm=1
      //   bk[1] → rc t="2" (XLDAPR) → cm=2
      const xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        '  xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">',
        '  <metadataTypes count="2">',
        '    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" cellMeta="1"/>',
        '    <metadataType name="XLDAPR" minSupportedVersion="120000" cellMeta="1"/>',
        "  </metadataTypes>",
        '  <futureMetadata name="XLDAPR" count="1">',
        '    <bk><extLst><ext><xda:dynamicArrayProperties fDynamic="1"/></ext></extLst></bk>',
        "  </futureMetadata>",
        '  <cellMetadata count="2">',
        "    <bk>",
        '      <rc t="1" v="0"/>',
        "    </bk>",
        "    <bk>",
        '      <rc t="2" v="0"/>',
        "    </bk>",
        "  </cellMetadata>",
        "</metadata>"
      ].join("\n");

      const xform = new MetadataXform();
      const result = await xform.parseStream(xmlToStream(xml));
      expect(result.hasDynamicArrays).toBe(true);
      // cm=1 maps to XLRICHVALUE → should NOT be in set
      expect(result.dynamicArrayCmIndices.has(1)).toBe(false);
      // cm=2 maps to XLDAPR → should be in set
      expect(result.dynamicArrayCmIndices.has(2)).toBe(true);
      expect(result.dynamicArrayCmIndices.size).toBe(1);
    });
  });
});
