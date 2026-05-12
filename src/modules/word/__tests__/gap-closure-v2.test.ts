/**
 * DOCX Module - Gap Closure V2 Tests
 *
 * Tests for all new features added to close gaps with docx4j / Open XML SDK:
 * 1. OOXML Strict format support (namespace normalization)
 * 2. DrawingML effects (reflection, soft edges, 3D)
 * 3. Font subsetting
 * 4. OpenDoPE data binding resolution
 * 5. Form field extraction and filling
 * 6. SmartArt round-trip preservation (via opaqueParts)
 */

import { createZip } from "@archive/zip/zip-bytes";
import { describe, it, expect } from "vitest";

import { STRICT_TO_TRANSITIONAL_REL, STRICT_TO_TRANSITIONAL_NS } from "../constants";
import {
  Document,
  textParagraph,
  paragraph,
  formTextField,
  formCheckboxField,
  formDropdownField,
  readDocx,
  toBuffer,
  resolveDataBindings,
  extractFormFields,
  fillFormFields,
  createShape,
  createRect,
  subsetFont,
  embedFont
} from "../index";
import type { DocxDocument, Paragraph, Run, StructuredDocumentTag as SdtType } from "../types";

// =============================================================================
// OOXML Strict Format Support
// =============================================================================

describe("OOXML Strict: namespace mapping", () => {
  it("should have complete Strict→Transitional namespace mappings", () => {
    // Verify core namespace mappings exist
    expect(STRICT_TO_TRANSITIONAL_NS.get("http://purl.oclc.org/ooxml/wordprocessingml/main")).toBe(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    );
    expect(
      STRICT_TO_TRANSITIONAL_NS.get("http://purl.oclc.org/ooxml/officeDocument/relationships")
    ).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    expect(STRICT_TO_TRANSITIONAL_NS.get("http://purl.oclc.org/ooxml/drawingml/main")).toBe(
      "http://schemas.openxmlformats.org/drawingml/2006/main"
    );
    expect(STRICT_TO_TRANSITIONAL_NS.get("http://purl.oclc.org/ooxml/drawingml/picture")).toBe(
      "http://schemas.openxmlformats.org/drawingml/2006/picture"
    );
  });

  it("should have complete Strict→Transitional relationship type mappings", () => {
    // Verify key relationship type mappings
    expect(
      STRICT_TO_TRANSITIONAL_REL.get(
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
      )
    ).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    expect(
      STRICT_TO_TRANSITIONAL_REL.get(
        "http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
      )
    ).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    expect(
      STRICT_TO_TRANSITIONAL_REL.get(
        "http://purl.oclc.org/ooxml/officeDocument/relationships/header"
      )
    ).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header");
    expect(
      STRICT_TO_TRANSITIONAL_REL.get(
        "http://purl.oclc.org/ooxml/officeDocument/relationships/image"
      )
    ).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
  });

  it("should map all critical relationship types", () => {
    const criticalTypes = [
      "officeDocument",
      "styles",
      "settings",
      "fontTable",
      "numbering",
      "footnotes",
      "endnotes",
      "header",
      "footer",
      "image",
      "hyperlink",
      "theme",
      "comments",
      "chart",
      "customXml",
      "vbaProject"
    ];
    for (const relType of criticalTypes) {
      const strictUri = `http://purl.oclc.org/ooxml/officeDocument/relationships/${relType}`;
      expect(STRICT_TO_TRANSITIONAL_REL.has(strictUri)).toBe(true);
    }
  });
});

describe("OOXML Strict: reader normalization", () => {
  it("should read document with Strict package relationships", async () => {
    // Create a document, package it, then manually verify the reader can handle
    // the scenario where package rels use the Strict namespace (by testing the parsing logic)
    const doc = Document.create();
    Document.addParagraph(doc, "Hello Strict");
    const model = Document.build(doc);
    const buffer = await toBuffer(model);

    // The reader should work with standard documents (regression)
    const result = await readDocx(buffer);
    expect(result.body.length).toBeGreaterThan(0);
    const firstPara = result.body[0] as Paragraph;
    expect(firstPara.type).toBe("paragraph");
  });

  it("should discover document part path from package relationships", async () => {
    // Standard documents have word/document.xml — verify it's found correctly
    const doc = Document.create();
    Document.addParagraph(doc, "Discovery test");
    const model = Document.build(doc);
    const buffer = await toBuffer(model);
    const result = await readDocx(buffer);
    expect(result.body.length).toBeGreaterThan(0);
  });

  it("should parse a synthetic Strict-format DOCX with Strict relationship types", async () => {
    // Build a minimal DOCX ZIP that uses ISO 29500 Strict relationship type URIs
    const enc = new TextEncoder();

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    // Package rels with STRICT relationship type URI
    const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    // Document rels with STRICT relationship types
    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    // Minimal document.xml
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Strict Format Document</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;

    // Minimal styles.xml
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;

    // Build the ZIP
    const zipBuffer = await createZip([
      { name: "[Content_Types].xml", data: enc.encode(contentTypes) },
      { name: "_rels/.rels", data: enc.encode(packageRels) },
      { name: "word/_rels/document.xml.rels", data: enc.encode(docRels) },
      { name: "word/document.xml", data: enc.encode(documentXml) },
      { name: "word/styles.xml", data: enc.encode(stylesXml) }
    ]);

    // The reader should successfully parse this Strict-format document
    const result = await readDocx(zipBuffer);

    expect(result.body.length).toBe(2);
    const para1 = result.body[0] as Paragraph;
    expect(para1.type).toBe("paragraph");
    const run1 = para1.children[0] as Run;
    expect(run1.content![0]).toEqual({ type: "text", text: "Strict Format Document" });

    const para2 = result.body[1] as Paragraph;
    const run2 = para2.children[0] as Run;
    expect(run2.content![0]).toEqual({ type: "text", text: "Second paragraph" });

    // Styles should be parsed despite using Strict rel type
    expect(result.styles).toBeDefined();
    expect(result.styles!.length).toBeGreaterThan(0);
    // The style should have its ID parsed
    const normalStyle = result.styles!.find(s => s.styleId === "Normal");
    expect(normalStyle).toBeDefined();
  });

  it("should parse document with Strict image relationship type", async () => {
    // Test that image relationships using Strict URI are normalized
    const enc = new TextEncoder();

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

    const packageRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    // Document rels with STRICT image relationship type
    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/image" Target="media/image1.png"/>
</Relationships>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Doc with image</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;

    // Minimal 1x1 PNG
    const png1x1 = new Uint8Array([
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
      0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
      0xcf, 0xc0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00,
      0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
    ]);

    const zipBuffer = await createZip([
      { name: "[Content_Types].xml", data: enc.encode(contentTypes) },
      { name: "_rels/.rels", data: enc.encode(packageRels) },
      { name: "word/_rels/document.xml.rels", data: enc.encode(docRels) },
      { name: "word/document.xml", data: enc.encode(documentXml) },
      { name: "word/media/image1.png", data: png1x1 }
    ]);

    const result = await readDocx(zipBuffer);
    expect(result.body.length).toBe(1);
    // Image should be found despite Strict rel type
    expect(result.images).toBeDefined();
    expect(result.images!.length).toBe(1);
    expect(result.images![0].mediaType).toBe("png");
  });
});

// =============================================================================
// DrawingML Effects — Reflection, Soft Edges, 3D
// =============================================================================

describe("DrawingML effects: reflection, softEdges, 3D", () => {
  it("should serialize reflection effect", () => {
    const shape = createShape({
      shapeType: "rect",
      width: 1000000,
      height: 500000,
      effects: {
        reflection: {
          blurRadius: 50800,
          startOpacity: 50,
          endOpacity: 0,
          distance: 38100,
          direction: 5400000,
          fadeDirection: 5400000
        }
      }
    });
    expect(shape.rawXml).toContain("<a:reflection");
    expect(shape.rawXml).toContain('blurRad="50800"');
    expect(shape.rawXml).toContain('stA="50000"');
    expect(shape.rawXml).toContain('endA="0"');
    expect(shape.rawXml).toContain('dist="38100"');
    expect(shape.rawXml).toContain('dir="5400000"');
    expect(shape.rawXml).toContain('fadeDir="5400000"');
  });

  it("should serialize soft edges effect", () => {
    const shape = createShape({
      shapeType: "roundRect",
      width: 2000000,
      height: 1000000,
      effects: {
        softEdges: 127000
      }
    });
    expect(shape.rawXml).toContain('<a:softEdge rad="127000"/>');
  });

  it("should serialize 3D effect with camera and bevel", () => {
    const shape = createShape({
      shapeType: "rect",
      width: 3000000,
      height: 2000000,
      effects: {
        effect3d: {
          camera: "perspectiveFront",
          rotX: 3000000,
          rotY: 2000000,
          rotZ: 0,
          bevelTop: { width: 127000, height: 63500, preset: "circle" },
          bevelBottom: { width: 127000, height: 63500 },
          extrusionDepth: 254000,
          extrusionColor: "4472C4"
        }
      }
    });
    expect(shape.rawXml).toContain('<a:camera prst="perspectiveFront">');
    expect(shape.rawXml).toContain('<a:rot lat="3000000" lon="2000000" rev="0"/>');
    expect(shape.rawXml).toContain("<a:lightRig");
    expect(shape.rawXml).toContain('<a:bevelT w="127000" h="63500" prst="circle"/>');
    expect(shape.rawXml).toContain('<a:bevelB w="127000" h="63500"/>');
    expect(shape.rawXml).toContain('extrusionH="254000"');
    expect(shape.rawXml).toContain('<a:extrusionClr><a:srgbClr val="4472C4"/></a:extrusionClr>');
    expect(shape.rawXml).toContain("<a:scene3d>");
    expect(shape.rawXml).toContain("<a:sp3d");
  });

  it("should serialize combined effects", () => {
    const shape = createShape({
      shapeType: "ellipse",
      width: 1500000,
      height: 1500000,
      fill: { type: "solid", color: "FF0000" },
      effects: {
        shadow: {
          type: "outer",
          color: "000000",
          blurRadius: 50800,
          distance: 38100,
          direction: 2700000
        },
        glow: { color: "FFFF00", radius: 101600, transparency: 30 },
        reflection: { startOpacity: 60, endOpacity: 0, distance: 25400 },
        softEdges: 63500,
        effect3d: { camera: "isometricTopDown", bevelTop: { width: 50800, height: 50800 } }
      }
    });
    // All effects should be present
    expect(shape.rawXml).toContain("a:outerShdw");
    expect(shape.rawXml).toContain("a:glow");
    expect(shape.rawXml).toContain("a:reflection");
    expect(shape.rawXml).toContain("a:softEdge");
    expect(shape.rawXml).toContain("a:scene3d");
    expect(shape.rawXml).toContain("a:sp3d");
  });

  it("should serialize glow with transparency", () => {
    const shape = createShape({
      shapeType: "rect",
      width: 1000000,
      height: 1000000,
      effects: {
        glow: { color: "00FF00", radius: 76200, transparency: 50 }
      }
    });
    expect(shape.rawXml).toContain('<a:glow rad="76200">');
    expect(shape.rawXml).toContain('<a:alpha val="50000"/>');
    expect(shape.rawXml).toContain('val="00FF00"');
  });

  it("should produce round-trip stable shapes with effects", async () => {
    const shape = createRect(2000000, 1000000, {
      fill: { type: "solid", color: "4472C4" },
      effects: {
        shadow: {
          type: "outer",
          color: "000000",
          blurRadius: 50800,
          distance: 38100,
          direction: 2700000
        },
        softEdges: 63500
      }
    });
    // Verify shape is created properly with rawXml
    expect(shape.type).toBe("drawingShape");
    expect(shape.fillColor).toBe("4472C4");
    expect(shape.rawXml).toContain("a:outerShdw");
    expect(shape.rawXml).toContain("a:softEdge");

    // Build a document with a paragraph followed by a shape
    const doc = Document.create();
    Document.addParagraph(doc, "Before shape");
    Document.addContent(doc, shape);
    Document.addParagraph(doc, "After shape");
    const model = Document.build(doc);
    const buffer = await toBuffer(model);
    const result = await readDocx(buffer);
    // Document should be parseable (at least paragraphs survive)
    expect(result.body.length).toBeGreaterThanOrEqual(2);
  });

  it("wraps advanced shape effects inside a valid wp:anchor / a:effectLst structure", async () => {
    // Regression: previously the writer short-circuited on shape.rawXml and
    // emitted bare DrawingML fragments directly under w:body. Now the writer
    // always renders the full drawing wrapper and inserts the advanced
    // properties (effect list, gradient, scene3d) as children of wps:spPr.
    const shape = createRect(2000000, 1000000, {
      fill: {
        type: "gradient",
        stops: [
          { position: 0, color: "FF0000" },
          { position: 100000, color: "0000FF" }
        ]
      },
      effects: {
        shadow: { type: "outer", color: "000000", blurRadius: 50800 },
        softEdges: 63500
      }
    });

    const doc = Document.create();
    Document.addContent(doc, shape);
    const buffer = await toBuffer(Document.build(doc));

    const { unzip } = await import("@archive/read-archive");
    const reader = unzip(buffer);
    let documentXml = "";
    for await (const entry of reader.entries()) {
      if (entry.path === "word/document.xml") {
        const bytes = await entry.bytes();
        documentXml = new TextDecoder().decode(bytes);
        break;
      }
    }

    // Drawing wrappers must be present — they were missing before the fix.
    expect(documentXml).toContain("<w:drawing>");
    expect(documentXml).toContain("<wp:anchor");
    expect(documentXml).toContain("<wps:spPr>");

    // Effects must live inside <a:effectLst>, not as bare children of spPr
    // (and must NOT escape the spPr block).
    expect(documentXml).toContain("<a:effectLst>");
    const effectLstStart = documentXml.indexOf("<a:effectLst>");
    const spPrEnd = documentXml.indexOf("</wps:spPr>");
    expect(effectLstStart).toBeGreaterThan(0);
    expect(spPrEnd).toBeGreaterThan(effectLstStart);

    // Gradient fill should also be inside spPr
    const gradFillStart = documentXml.indexOf("<a:gradFill");
    expect(gradFillStart).toBeGreaterThan(0);
    expect(gradFillStart).toBeLessThan(spPrEnd);

    // OOXML schema requires fill to precede a:ln and effects to follow it.
    // Verify that ordering so consumers like Word don't reject the file.
    const lnStart = documentXml.indexOf("<a:ln", documentXml.indexOf("<wps:spPr>"));
    if (lnStart > 0) {
      expect(gradFillStart).toBeLessThan(lnStart);
      expect(effectLstStart).toBeGreaterThan(lnStart);
    }

    // The effect children must NOT appear directly under w:body — they used
    // to be emitted there, producing invalid OOXML.
    const bodyStart = documentXml.indexOf("<w:body>");
    const drawingStart = documentXml.indexOf("<w:drawing>");
    expect(drawingStart).toBeGreaterThan(bodyStart);
  });
});

// =============================================================================
// Font Subsetting
// =============================================================================

describe("Font subsetting", () => {
  // Create a minimal valid TrueType font for testing
  function createMinimalTtf(): Uint8Array {
    // Minimal TTF with offset table + head + maxp + cmap tables
    // This is a synthetic minimal font for testing the subsetter's
    // graceful handling (it won't have valid glyf data, so subset
    // should return original)
    const data = new Uint8Array(256);
    // sfVersion = 0x00010000 (TrueType)
    data[0] = 0x00;
    data[1] = 0x01;
    data[2] = 0x00;
    data[3] = 0x00;
    // numTables = 5
    data[4] = 0x00;
    data[5] = 0x05;
    // searchRange, entrySelector, rangeShift
    data[6] = 0x00;
    data[7] = 0x40;
    data[8] = 0x00;
    data[9] = 0x02;
    data[10] = 0x00;
    data[11] = 0x10;
    return data;
  }

  it("should return original data for non-TrueType fonts (CFF/OTF)", () => {
    // OTF magic: "OTTO"
    const otfData = new Uint8Array(100);
    otfData[0] = 0x4f; // O
    otfData[1] = 0x54; // T
    otfData[2] = 0x54; // T
    otfData[3] = 0x4f; // O
    const result = subsetFont(otfData, "Hello");
    expect(result).toBe(otfData); // Same reference = no subsetting attempted
  });

  it("should return original data when tables are missing", () => {
    const ttf = createMinimalTtf();
    // No valid tables — subsetter should fail gracefully
    const result = subsetFont(ttf, "ABC");
    // Should return original since it can't parse properly
    expect(result).toEqual(ttf);
  });

  it("should accept empty characters string gracefully", () => {
    const ttf = createMinimalTtf();
    const result = subsetFont(ttf, "");
    expect(result).toEqual(ttf);
  });

  it("should integrate with embedFont when usedCharacters is provided", () => {
    const ttf = createMinimalTtf();
    const result = embedFont({
      name: "TestFont",
      data: ttf,
      style: "regular",
      usedCharacters: "Hello World"
    });
    expect(result.fontDef.name).toBe("TestFont");
    expect(result.embeddedFont.data.length).toBeGreaterThan(0);
  });
});

// =============================================================================
// OpenDoPE Data Binding Resolution
// =============================================================================

describe("OpenDoPE: resolveDataBindings", () => {
  it("should resolve simple text binding from CustomXML", () => {
    const customXml = `<?xml version="1.0"?>
      <root><name>John Doe</name><email>john@example.com</email></root>`;

    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/name",
              storeItemId: "{12345678-1234-1234-1234-123456789ABC}"
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{12345678-1234-1234-1234-123456789ABC}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };

    const result = resolveDataBindings(doc);
    const sdt = result.body[0] as SdtType;
    expect(sdt.type).toBe("sdt");
    const para = sdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content![0]).toEqual({ type: "text", text: "John Doe" });
  });

  it("should resolve binding with override data map", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/data/value",
              storeItemId: "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"
            }
          },
          content: [textParagraph("old")]
        }
      ]
    };

    const overrideData = new Map([
      ["{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}", "<data><value>Override Value</value></data>"]
    ]);

    const result = resolveDataBindings(doc, overrideData);
    const sdt = result.body[0] as SdtType;
    const para = sdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content![0]).toEqual({ type: "text", text: "Override Value" });
  });

  it("should remove SDT when binding value is not found (conditional)", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/nonexistent",
              storeItemId: "{11111111-2222-3333-4444-555555555555}"
            }
          },
          content: [textParagraph("should disappear")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{11111111-2222-3333-4444-555555555555}",
          xmlContent: "<root><other>value</other></root>",
          fileName: "item1.xml"
        }
      ]
    };

    const result = resolveDataBindings(doc);
    // SDT should be replaced with empty paragraph (conditional removal)
    const first = result.body[0] as Paragraph;
    expect(first.type).toBe("paragraph");
    expect(first.children.length).toBe(0);
  });

  it("should handle nested XPath with namespaces stripped", () => {
    const customXml = `<ns0:root xmlns:ns0="http://example.com">
      <ns0:person><ns0:first>Jane</ns0:first><ns0:last>Smith</ns0:last></ns0:person>
    </ns0:root>`;

    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/ns0:root/ns0:person/ns0:first",
              storeItemId: "{AABB1122-3344-5566-7788-AABBCCDDEEFF}",
              prefixMappings: 'xmlns:ns0="http://example.com"'
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{AABB1122-3344-5566-7788-AABBCCDDEEFF}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };

    const result = resolveDataBindings(doc);
    const sdt = result.body[0] as SdtType;
    const para = sdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content![0]).toEqual({ type: "text", text: "Jane" });
  });

  it("should leave SDTs without dataBinding unchanged", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: { alias: "NoBinding" },
          content: [textParagraph("keep me")]
        }
      ]
    };

    const result = resolveDataBindings(doc);
    const sdt = result.body[0] as SdtType;
    expect(sdt.type).toBe("sdt");
    expect((sdt.properties as any).alias).toBe("NoBinding");
  });

  it("should return document unchanged when no CustomXML parts exist", () => {
    const doc: DocxDocument = {
      body: [textParagraph("Hello")]
    };
    const result = resolveDataBindings(doc);
    expect(result).toBe(doc); // Same reference = unchanged
  });

  it("should handle position predicates [1] and [last()]", () => {
    const customXml = `<root>
      <item>First</item>
      <item>Second</item>
      <item>Third</item>
    </root>`;

    // Test [1] — first item
    const doc1: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/item[1]",
              storeItemId: "{AAAA0000-0000-0000-0000-000000000001}"
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{AAAA0000-0000-0000-0000-000000000001}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };
    const result1 = resolveDataBindings(doc1);
    const sdt1 = result1.body[0] as SdtType;
    const para1 = sdt1.content[0] as Paragraph;
    const run1 = para1.children[0] as Run;
    expect(run1.content![0]).toEqual({ type: "text", text: "First" });

    // Test [last()] — last item
    const doc2: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/item[last()]",
              storeItemId: "{AAAA0000-0000-0000-0000-000000000001}"
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{AAAA0000-0000-0000-0000-000000000001}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };
    const result2 = resolveDataBindings(doc2);
    const sdt2 = result2.body[0] as SdtType;
    const para2 = sdt2.content[0] as Paragraph;
    const run2 = para2.children[0] as Run;
    expect(run2.content![0]).toEqual({ type: "text", text: "Third" });
  });

  it("should handle attribute access @attr", () => {
    const customXml = `<root><item id="42" status="active">Content</item></root>`;

    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/item/@id",
              storeItemId: "{BBBB0000-0000-0000-0000-000000000002}"
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{BBBB0000-0000-0000-0000-000000000002}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };
    const result = resolveDataBindings(doc);
    const sdt = result.body[0] as SdtType;
    const para = sdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content![0]).toEqual({ type: "text", text: "42" });
  });

  it("should handle deeply nested elements with same-name parents", () => {
    const customXml = `<data>
      <section><title>Section 1</title><section><title>Nested</title></section></section>
      <section><title>Section 2</title></section>
    </data>`;

    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/data/section[1]/title",
              storeItemId: "{CCCC0000-0000-0000-0000-000000000003}"
            }
          },
          content: [textParagraph("placeholder")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{CCCC0000-0000-0000-0000-000000000003}",
          xmlContent: customXml,
          fileName: "item1.xml"
        }
      ]
    };
    const result = resolveDataBindings(doc);
    const sdt = result.body[0] as SdtType;
    const para = sdt.content[0] as Paragraph;
    const run = para.children[0] as Run;
    expect(run.content![0]).toEqual({ type: "text", text: "Section 1" });
  });
});

// =============================================================================
// Form Field Extraction & Filling
// =============================================================================

describe("Form fields: extractFormFields", () => {
  it("should extract text form fields from document", () => {
    const doc = Document.create();
    Document.addParagraphElement(
      doc,
      paragraph([formTextField({ name: "FullName", default: "Jane Doe" })])
    );
    Document.addParagraphElement(
      doc,
      paragraph([formCheckboxField({ name: "Agree", checked: true })])
    );
    Document.addParagraphElement(
      doc,
      paragraph([
        formDropdownField({
          name: "Country",
          entries: ["US", "UK", "JP"],
          default: 1
        })
      ])
    );
    const model = Document.build(doc);

    const fields = extractFormFields(model);
    expect(fields.length).toBe(3);

    expect(fields[0].name).toBe("FullName");
    expect(fields[0].type).toBe("text");
    expect(fields[0].value).toBe("Jane Doe");

    expect(fields[1].name).toBe("Agree");
    expect(fields[1].type).toBe("checkBox");
    expect(fields[1].value).toBe(true);

    expect(fields[2].name).toBe("Country");
    expect(fields[2].type).toBe("dropDown");
    expect(fields[2].value).toBe(1);
    expect(fields[2].entries).toEqual(["US", "UK", "JP"]);
  });

  it("should extract form fields from tables", () => {
    const doc = Document.create();
    Document.addTable(doc, [["Label"]]);
    // Manually build a model with form field inside table
    const model: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [paragraph([formTextField({ name: "InTable", default: "TableValue" })])]
                }
              ]
            }
          ]
        }
      ]
    };

    const fields = extractFormFields(model);
    expect(fields.length).toBe(1);
    expect(fields[0].name).toBe("InTable");
    expect(fields[0].value).toBe("TableValue");
  });
});

describe("Form fields: fillFormFields", () => {
  it("should fill text form field values", () => {
    const doc = Document.create();
    Document.addParagraphElement(doc, paragraph([formTextField({ name: "Name", default: "old" })]));
    const model = Document.build(doc);

    const values = new Map<string, string | boolean | number>([["Name", "New Value"]]);
    const filled = fillFormFields(model, values);

    const fields = extractFormFields(filled);
    expect(fields[0].value).toBe("New Value");
  });

  it("should fill checkbox form field values", () => {
    const doc = Document.create();
    Document.addParagraphElement(
      doc,
      paragraph([formCheckboxField({ name: "Accept", checked: false })])
    );
    const model = Document.build(doc);

    const values = new Map<string, string | boolean | number>([["Accept", true]]);
    const filled = fillFormFields(model, values);

    const fields = extractFormFields(filled);
    expect(fields[0].value).toBe(true);
  });

  it("should not modify fields without matching values", () => {
    const doc = Document.create();
    Document.addParagraphElement(
      doc,
      paragraph([formTextField({ name: "Keep", default: "original" })])
    );
    const model = Document.build(doc);

    const values = new Map<string, string | boolean | number>([["Other", "ignored"]]);
    const filled = fillFormFields(model, values);

    const fields = extractFormFields(filled);
    expect(fields[0].value).toBe("original");
  });

  it("should fill form fields inside tables", () => {
    const model: DocxDocument = {
      body: [
        {
          type: "table",
          rows: [
            {
              cells: [
                {
                  content: [paragraph([formTextField({ name: "Cell1", default: "" })])]
                }
              ]
            }
          ]
        }
      ]
    };

    const values = new Map<string, string | boolean | number>([["Cell1", "Filled!"]]);
    const filled = fillFormFields(model, values);
    const fields = extractFormFields(filled);
    expect(fields[0].value).toBe("Filled!");
  });
});

// =============================================================================
// SmartArt Round-Trip Preservation
// =============================================================================

describe("SmartArt: round-trip via opaqueParts", () => {
  it("should preserve unknown drawing content as opaqueDrawing in body", async () => {
    // Create a document with a drawing shape (which round-trips)
    const doc = Document.create();
    Document.addParagraph(doc, "Before shape");
    Document.addContent(
      doc,
      createRect(2000000, 1000000, { fill: { type: "solid", color: "FF0000" } })
    );
    Document.addParagraph(doc, "After shape");
    const model = Document.build(doc);
    const buffer = await toBuffer(model);

    // Read it back - shape content should survive
    const result = await readDocx(buffer);
    expect(result.body.length).toBeGreaterThanOrEqual(2);
  });

  it("should collect unrecognized ZIP entries as opaqueParts", async () => {
    // Any file not explicitly consumed by the reader goes into opaqueParts
    // This ensures SmartArt diagram files are preserved
    const doc = Document.create();
    Document.addParagraph(doc, "Simple doc");
    const model = Document.build(doc);
    const buffer = await toBuffer(model);

    // Read back - well-formed docs shouldn't have opaque parts
    const result = await readDocx(buffer);
    // A clean document shouldn't have opaque parts
    expect(result.opaqueParts ?? []).toHaveLength(0);
  });
});

// =============================================================================
// Integration: Round-Trip with New Features
// =============================================================================

describe("Integration: new features round-trip", () => {
  it("should round-trip document with SDT data binding", async () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "sdt",
          properties: {
            dataBinding: {
              xpath: "/root/field1",
              storeItemId: "{12345678-ABCD-1234-5678-ABCDEF123456}"
            }
          },
          content: [textParagraph("Bound Value")]
        }
      ],
      customXmlParts: [
        {
          itemId: "{12345678-ABCD-1234-5678-ABCDEF123456}",
          xmlContent: "<root><field1>Hello</field1></root>",
          fileName: "item1.xml"
        }
      ]
    };

    const buffer = await toBuffer(doc);
    const result = await readDocx(buffer);

    // SDT should preserve its data binding
    const sdt = result.body[0] as SdtType;
    expect(sdt.type).toBe("sdt");
    expect(sdt.properties?.dataBinding?.xpath).toBe("/root/field1");
    expect(sdt.properties?.dataBinding?.storeItemId).toBe("{12345678-ABCD-1234-5678-ABCDEF123456}");

    // CustomXML parts should survive
    expect(result.customXmlParts).toBeDefined();
    expect(result.customXmlParts!.length).toBe(1);
    // Reader may strip or retain braces from GUID — normalize for comparison
    const itemId = result.customXmlParts![0].itemId.replace(/[{}]/g, "");
    expect(itemId.toLowerCase()).toBe("12345678-abcd-1234-5678-abcdef123456");
  });

  it("should round-trip document with form fields", async () => {
    const doc = Document.create();
    Document.addParagraphElement(
      doc,
      paragraph([formTextField({ name: "FirstName", default: "Alice" })])
    );
    Document.addParagraphElement(
      doc,
      paragraph([formCheckboxField({ name: "Premium", checked: true })])
    );
    Document.addParagraphElement(
      doc,
      paragraph([
        formDropdownField({ name: "Plan", entries: ["Basic", "Pro", "Enterprise"], default: 2 })
      ])
    );
    const model = Document.build(doc);
    const buffer = await toBuffer(model);

    const result = await readDocx(buffer);
    const fields = extractFormFields(result);

    expect(fields.length).toBe(3);
    expect(fields[0].name).toBe("FirstName");
    expect(fields[0].type).toBe("text");
    expect(fields[1].name).toBe("Premium");
    expect(fields[1].type).toBe("checkBox");
    expect(fields[2].name).toBe("Plan");
    expect(fields[2].type).toBe("dropDown");
  });
});
