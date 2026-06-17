/**
 * DOCX XML Compliance Tests
 *
 * Validates Open XML structural compliance by generating DOCX packages,
 * extracting them, parsing the XML parts, and verifying element structure
 * using DOM-based assertions instead of string matching.
 */

import { extractAll } from "@archive/unzip/extract";
import { Document, Build, Io } from "@word/index";
import type { Table } from "@word/index";
import { parseXml, findChild, findChildren, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";
import { describe, it, expect } from "vitest";

// Minimal 1x1 PNG for image tests
const MINI_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
  0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82
]);

const decoder = new TextDecoder();

/** Helper: extract all files from DOCX bytes into a Map<path, Uint8Array>. */
async function extractDocx(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const files = await extractAll(bytes);
  const result = new Map<string, Uint8Array>();
  for (const [path, entry] of files) {
    result.set(path, entry.data);
  }
  return result;
}

/** Helper: decode and parse an XML entry from the ZIP. */
function parseEntry(files: Map<string, Uint8Array>, path: string): XmlElement {
  const data = files.get(path);
  if (!data) {
    throw new Error(`Entry not found in ZIP: ${path}`);
  }
  const doc = parseXml(decoder.decode(data));
  return doc.root;
}

/** Recursively find the first descendant element matching a name. */
function findDescendant(el: XmlElement, name: string): XmlElement | undefined {
  for (const child of el.children) {
    if (child.type === "element") {
      if (child.name === name) {
        return child;
      }
      const found = findDescendant(child, name);
      if (found) {
        return found;
      }
    }
  }
  return undefined;
}

/** Recursively find all descendant elements matching a name. */
function findDescendants(el: XmlElement, name: string): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of el.children) {
    if (child.type === "element") {
      if (child.name === name) {
        result.push(child);
      }
      result.push(...findDescendants(child, name));
    }
  }
  return result;
}

// =============================================================================
// Tests
// =============================================================================

describe("DOCX XML Compliance", () => {
  // ===========================================================================
  // 1. document.xml 结构合规
  // ===========================================================================

  describe("document.xml structure", () => {
    it("should have w:document root with namespace, w:body, w:p, w:r, w:t elements", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Hello World");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");

      // Root element is w:document
      expect(root.name).toBe("w:document");

      // Has xmlns:w namespace declaration
      const nsW = root.attributes["xmlns:w"] ?? (root.ns ? root.ns["w"] : undefined);
      expect(nsW).toBeDefined();

      // Has w:body child
      const body = findChild(root, "w:body");
      expect(body).toBeDefined();

      // Body has w:p paragraph
      const paragraphs = findChildren(body!, "w:p");
      expect(paragraphs.length).toBeGreaterThan(0);

      // Paragraph has w:r run
      const runs = findChildren(paragraphs[0], "w:r");
      expect(runs.length).toBeGreaterThan(0);

      // Run has w:t text element
      const wt = findChild(runs[0], "w:t");
      expect(wt).toBeDefined();
      expect(textContent(wt!)).toBe("Hello World");
    });
  });

  // ===========================================================================
  // 2. styles.xml 结构合规
  // ===========================================================================

  describe("styles.xml structure", () => {
    it("should have w:styles root with w:docDefaults and w:style children", async () => {
      const h = Document.create();
      Document.useDefaultStyles(h);
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/styles.xml");

      // Root is w:styles
      expect(root.name).toBe("w:styles");

      // Has w:docDefaults
      const docDefaults = findChild(root, "w:docDefaults");
      expect(docDefaults).toBeDefined();

      // Has w:style children (at least Normal)
      const styles = findChildren(root, "w:style");
      expect(styles.length).toBeGreaterThan(0);

      // Each style has w:type and w:styleId attributes
      for (const style of styles) {
        expect(style.attributes["w:type"]).toBeDefined();
        expect(style.attributes["w:styleId"]).toBeDefined();
      }

      // Normal style exists
      const normal = styles.find(s => s.attributes["w:styleId"] === "Normal");
      expect(normal).toBeDefined();
    });
  });

  // ===========================================================================
  // 3. numbering.xml 结构合规
  // ===========================================================================

  describe("numbering.xml structure", () => {
    it("should have w:numbering root with abstractNum and num elements when lists exist", async () => {
      const h = Document.create();
      Document.addBulletList(h, ["Item 1", "Item 2"]);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/numbering.xml");

      // Root is w:numbering
      expect(root.name).toBe("w:numbering");

      // Has w:abstractNum children
      const abstractNums = findChildren(root, "w:abstractNum");
      expect(abstractNums.length).toBeGreaterThan(0);

      // Each abstractNum has w:abstractNumId attribute and w:lvl children
      for (const absNum of abstractNums) {
        expect(absNum.attributes["w:abstractNumId"]).toBeDefined();
        const levels = findChildren(absNum, "w:lvl");
        expect(levels.length).toBeGreaterThan(0);
      }

      // Has w:num children referencing abstractNumId
      const nums = findChildren(root, "w:num");
      expect(nums.length).toBeGreaterThan(0);

      for (const num of nums) {
        expect(num.attributes["w:numId"]).toBeDefined();
        const absNumIdRef = findChild(num, "w:abstractNumId");
        expect(absNumIdRef).toBeDefined();
        expect(absNumIdRef!.attributes["w:val"]).toBeDefined();
      }
    });
  });

  // ===========================================================================
  // 4. relationships 一致性
  // ===========================================================================

  describe("relationships consistency", () => {
    it("should have valid _rels/.rels with Id, Type, Target attributes", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "_rels/.rels");
      const rels = findChildren(root, "Relationship");
      expect(rels.length).toBeGreaterThan(0);

      for (const rel of rels) {
        expect(rel.attributes["Id"]).toBeDefined();
        expect(rel.attributes["Type"]).toBeDefined();
        expect(rel.attributes["Target"]).toBeDefined();
      }
    });

    it("should have document.xml.rels targets that exist in the ZIP", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/_rels/document.xml.rels");
      const rels = findChildren(root, "Relationship");

      for (const rel of rels) {
        const targetMode = rel.attributes["TargetMode"];
        // Skip external targets (e.g. hyperlinks)
        if (targetMode === "External") {
          continue;
        }

        const target = rel.attributes["Target"];
        expect(target).toBeDefined();

        // Resolve target relative to word/
        const resolvedPath = target!.startsWith("/") ? target!.substring(1) : `word/${target}`;

        expect(
          files.has(resolvedPath),
          `Target "${resolvedPath}" referenced in document.xml.rels should exist in ZIP`
        ).toBe(true);
      }
    });
  });

  // ===========================================================================
  // 5. content types 一致性
  // ===========================================================================

  describe("content types consistency", () => {
    it("should have Override entries whose PartName files exist in ZIP", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "[Content_Types].xml");
      const overrides = findChildren(root, "Override");

      for (const override of overrides) {
        const partName = override.attributes["PartName"];
        expect(partName).toBeDefined();
        // PartName starts with "/"
        const zipPath = partName!.startsWith("/") ? partName!.substring(1) : partName!;
        expect(
          files.has(zipPath),
          `PartName "${partName}" should correspond to an actual file in ZIP`
        ).toBe(true);
      }
    });

    it("should have Default entries covering rels and xml extensions", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "[Content_Types].xml");
      const defaults = findChildren(root, "Default");

      const extensions = defaults.map(d => d.attributes["Extension"]);
      expect(extensions).toContain("rels");
      expect(extensions).toContain("xml");
    });
  });

  // ===========================================================================
  // 6. 段落属性写入合规
  // ===========================================================================

  describe("paragraph properties compliance", () => {
    it("should write w:pStyle for heading paragraphs", async () => {
      const h = Document.create();
      Document.useDefaultStyles(h);
      Document.addHeading(h, "My Heading", 1);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;
      const paras = findChildren(body, "w:p");

      // Find the heading paragraph
      const headingPara = paras.find(p => {
        const pPr = findChild(p, "w:pPr");
        const pStyle = pPr ? findChild(pPr, "w:pStyle") : undefined;
        return pStyle?.attributes["w:val"] === "Heading1";
      });
      expect(headingPara).toBeDefined();
    });

    it("should write w:numPr with w:numId and w:ilvl for numbered paragraphs", async () => {
      const h = Document.create();
      Document.addBulletList(h, ["Item"]);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;
      const paras = findChildren(body, "w:p");

      // Find paragraph with numbering
      const numberedPara = paras.find(p => {
        const pPr = findChild(p, "w:pPr");
        return pPr ? findChild(pPr, "w:numPr") !== undefined : false;
      });
      expect(numberedPara).toBeDefined();

      const pPr = findChild(numberedPara!, "w:pPr")!;
      const numPr = findChild(pPr, "w:numPr")!;
      const numId = findChild(numPr, "w:numId");
      const ilvl = findChild(numPr, "w:ilvl");

      expect(numId).toBeDefined();
      expect(numId!.attributes["w:val"]).toBeDefined();
      expect(ilvl).toBeDefined();
      expect(ilvl!.attributes["w:val"]).toBeDefined();
    });

    it("should write w:jc for paragraph alignment", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Centered", { alignment: "center" });
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;
      const paras = findChildren(body, "w:p");

      const centeredPara = paras.find(p => {
        const pPr = findChild(p, "w:pPr");
        const jc = pPr ? findChild(pPr, "w:jc") : undefined;
        return jc?.attributes["w:val"] === "center";
      });
      expect(centeredPara).toBeDefined();
    });
  });

  // ===========================================================================
  // 7. Run 属性写入合规
  // ===========================================================================

  describe("run properties compliance", () => {
    it("should write w:b for bold runs", async () => {
      const h = Document.create();
      Document.addContent(h, Build.paragraph([Build.bold("Bold Text")]));
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const runs = findDescendants(root, "w:r");

      const boldRun = runs.find(r => {
        const rPr = findChild(r, "w:rPr");
        return rPr ? findChild(rPr, "w:b") !== undefined : false;
      });
      expect(boldRun).toBeDefined();
    });

    it("should write w:i for italic runs", async () => {
      const h = Document.create();
      Document.addContent(h, Build.paragraph([Build.italic("Italic Text")]));
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const runs = findDescendants(root, "w:r");

      const italicRun = runs.find(r => {
        const rPr = findChild(r, "w:rPr");
        return rPr ? findChild(rPr, "w:i") !== undefined : false;
      });
      expect(italicRun).toBeDefined();
    });

    it("should write w:sz with half-point value for font size", async () => {
      const h = Document.create();
      // size 24 half-points = 12pt
      Document.addContent(h, Build.paragraph([Build.text("Sized", { size: 24 })]));
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const runs = findDescendants(root, "w:r");

      const sizedRun = runs.find(r => {
        const rPr = findChild(r, "w:rPr");
        if (!rPr) {
          return false;
        }
        const sz = findChild(rPr, "w:sz");
        return sz?.attributes["w:val"] === "24";
      });
      expect(sizedRun).toBeDefined();
    });

    it("should write w:color with 6-digit hex value", async () => {
      const h = Document.create();
      Document.addContent(h, Build.paragraph([Build.text("Red Text", { color: "FF0000" })]));
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const runs = findDescendants(root, "w:r");

      const coloredRun = runs.find(r => {
        const rPr = findChild(r, "w:rPr");
        if (!rPr) {
          return false;
        }
        const color = findChild(rPr, "w:color");
        return color?.attributes["w:val"] === "FF0000";
      });
      expect(coloredRun).toBeDefined();
    });
  });

  // ===========================================================================
  // 8. 表格结构合规
  // ===========================================================================

  describe("table structure compliance", () => {
    it("should produce w:tbl with w:tblPr, w:tblGrid, w:tr, w:tc, and w:p", async () => {
      const h = Document.create();
      Document.addTable(
        h,
        [
          ["A", "B"],
          ["C", "D"]
        ],
        { columnWidths: [4000, 4000] }
      );
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;

      // Table is w:tbl
      const tbl = findChild(body, "w:tbl");
      expect(tbl).toBeDefined();

      // Has w:tblPr
      const tblPr = findChild(tbl!, "w:tblPr");
      expect(tblPr).toBeDefined();

      // Has w:tblGrid with w:gridCol children
      const tblGrid = findChild(tbl!, "w:tblGrid");
      expect(tblGrid).toBeDefined();
      const gridCols = findChildren(tblGrid!, "w:gridCol");
      expect(gridCols.length).toBe(2);

      // Has w:tr rows
      const rows = findChildren(tbl!, "w:tr");
      expect(rows.length).toBe(2);

      // First row has w:tc cells
      const cells = findChildren(rows[0], "w:tc");
      expect(cells.length).toBe(2);

      // Each cell has at least one w:p
      for (const tc of cells) {
        const paras = findChildren(tc, "w:p");
        expect(paras.length).toBeGreaterThan(0);
      }
    });

    it("should auto-synthesize w:tblGrid when columnWidths is omitted", async () => {
      // ECMA-376 §17.4.49 requires <w:tblGrid> with at least one
      // <w:gridCol>. Word rejects packages where the element is missing.
      const h = Document.create();
      Document.addTable(h, [
        ["A", "B", "C"],
        ["1", "2", "3"]
      ]);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);
      const tbl = findChild(findChild(parseEntry(files, "word/document.xml"), "w:body")!, "w:tbl")!;
      const tblGrid = findChild(tbl, "w:tblGrid");
      expect(tblGrid).toBeDefined();
      const gridCols = findChildren(tblGrid!, "w:gridCol");
      expect(gridCols.length).toBe(3);
      // All gridCol widths must be positive — Word treats w="0" as
      // invalid in dxa context.
      for (const c of gridCols) {
        expect(parseInt(c.attributes["w:w"]!, 10)).toBeGreaterThan(0);
      }
    });

    it("should ensure every w:tc emits a <w:tcPr><w:tcW/>", async () => {
      // Word strict mode rejects table cells without <w:tcW>.
      const h = Document.create();
      Document.addTable(h, [["only"]]);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);
      const tbl = findChild(findChild(parseEntry(files, "word/document.xml"), "w:body")!, "w:tbl")!;
      const tc = findChild(findChild(tbl, "w:tr")!, "w:tc")!;
      const tcPr = findChild(tc, "w:tcPr");
      expect(tcPr).toBeDefined();
      expect(findChild(tcPr!, "w:tcW")).toBeDefined();
    });

    it("should append a trailing <w:p> when a cell ends with a nested table", async () => {
      // ECMA-376 §17.4.66 requires every CT_Tc to end with a <w:p>.
      const inner: Table = {
        type: "table",
        properties: { width: { value: 5000, type: "pct" } },
        rows: [{ cells: [{ content: [{ type: "paragraph", children: [] }] }] }]
      };
      const outer: Table = {
        type: "table",
        properties: { width: { value: 5000, type: "pct" } },
        rows: [{ cells: [{ content: [inner] }] }]
      };
      const h = Document.create();
      Document.addTableElement(h, outer);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);
      const outerTbl = findChild(
        findChild(parseEntry(files, "word/document.xml"), "w:body")!,
        "w:tbl"
      )!;
      const outerCell = findChild(findChild(outerTbl, "w:tr")!, "w:tc")!;
      const outerCellChildren = outerCell.children.filter(c => c.type === "element");
      // Last child of every cell must be <w:p>
      const lastChild = outerCellChildren[outerCellChildren.length - 1];
      expect(lastChild.type).toBe("element");
      if (lastChild.type === "element") {
        expect(lastChild.name).toBe("w:p");
      }
    });

    it("should insert a separator <w:p> between two adjacent <w:tbl> blocks at body level", async () => {
      // ECMA-376 §17.13.5.34: a <w:tbl> must be followed by a paragraph
      // (or a section break) before the next <w:tbl> may appear.
      const h = Document.create();
      Document.addTable(h, [["A"]]);
      Document.addTable(h, [["B"]]);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);
      const body = findChild(parseEntry(files, "word/document.xml"), "w:body")!;
      const elementChildren = body.children.filter(c => c.type === "element");
      const elementNames = elementChildren.map(c => (c.type === "element" ? c.name : ""));
      // Walk and assert no two consecutive <w:tbl> exist (sectPr at the
      // tail is fine).
      for (let i = 1; i < elementNames.length; i++) {
        if (elementNames[i] === "w:tbl") {
          expect(elementNames[i - 1]).not.toBe("w:tbl");
        }
      }
    });
  });

  // ===========================================================================
  // CT_Body terminal-element rules — body must end with <w:p> or <w:sectPr>,
  // and headers/footers must contain at least one <w:p>.
  // ===========================================================================

  describe("body terminal-element compliance", () => {
    it("emits a default <w:sectPr> when the model omits sectionProperties", async () => {
      const doc = {
        body: [{ type: "paragraph", children: [] }]
      } as any;
      const bytes = await Io.package(doc);
      const files = await extractDocx(bytes);
      const body = findChild(parseEntry(files, "word/document.xml"), "w:body")!;
      const sectPr = findChild(body, "w:sectPr");
      expect(sectPr).toBeDefined();
      // Default page size must be present
      expect(findChild(sectPr!, "w:pgSz")).toBeDefined();
    });

    it("synthesises a <w:p> when the model body is empty", async () => {
      const doc = { body: [] } as any;
      const bytes = await Io.package(doc);
      const files = await extractDocx(bytes);
      const body = findChild(parseEntry(files, "word/document.xml"), "w:body")!;
      // Must contain at least one <w:p>
      const ps = findChildren(body, "w:p");
      expect(ps.length).toBeGreaterThan(0);
    });

    it("synthesises a <w:p> in headers/footers with empty children", async () => {
      const h = Document.create();
      Document.useDefaultStyles(h);
      Document.addParagraph(h, "body");
      Document.setHeader(h, "default", { children: [] });
      Document.setFooter(h, "default", { children: [] });
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);
      const headerName = [...files.keys()].find(n => /^word\/header\d+\.xml$/.test(n))!;
      const footerName = [...files.keys()].find(n => /^word\/footer\d+\.xml$/.test(n))!;
      const headerXml = decoder.decode(files.get(headerName)!);
      const footerXml = decoder.decode(files.get(footerName)!);
      // Both must contain at least one <w:p>
      expect(/<w:p[\s/>]/.test(headerXml)).toBe(true);
      expect(/<w:p[\s/>]/.test(footerXml)).toBe(true);
    });
  });

  // ===========================================================================
  // Field cached values must respect the same <w:t> newline rule.
  // ===========================================================================

  describe("field cached value compliance", () => {
    it("splits cachedValue containing \\n into <w:t>+<w:br/>", async () => {
      // Regression: index/TOC fields cached multi-line strings as
      // `<w:t>line1\nline2</w:t>` which Word rejects.
      const doc = {
        body: [
          {
            type: "paragraph",
            children: [
              {
                content: [
                  {
                    type: "field",
                    instruction: " INDEX ",
                    cachedValue: "alpha\tp1\nbeta\tp2"
                  }
                ]
              }
            ]
          }
        ]
      } as any;
      const bytes = await Io.package(doc);
      const files = await extractDocx(bytes);
      const docXml = decoder.decode(files.get("word/document.xml")!);
      // Cached value must NOT contain a literal newline inside a <w:t>
      const wtMatches = docXml.match(/<w:t[^>]*>[^<]*<\/w:t>/g) ?? [];
      for (const m of wtMatches) {
        expect(/[\r\n]/.test(m)).toBe(false);
      }
      // …and Word's required <w:br/> must appear once between the segments
      expect(docXml).toContain("alpha\tp1");
      expect(docXml).toContain("beta\tp2");
      expect(docXml).toMatch(/<w:br\s*\/>/);
    });
  });

  // ===========================================================================
  // 9. 超链接结构合规
  // ===========================================================================

  describe("hyperlink structure compliance", () => {
    it("should produce w:hyperlink with r:id and corresponding External relationship", async () => {
      const h = Document.create();
      Document.addContent(
        h,
        Build.paragraph([Build.hyperlink("Click me", { url: "https://example.com" })])
      );
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;

      // Find w:hyperlink in paragraph
      const hyps = findDescendants(body, "w:hyperlink");
      expect(hyps.length).toBeGreaterThan(0);

      const hyp = hyps[0];
      const rId = hyp.attributes["r:id"];
      expect(rId).toBeDefined();

      // Verify the relationship exists in document.xml.rels
      const relsRoot = parseEntry(files, "word/_rels/document.xml.rels");
      const rels = findChildren(relsRoot, "Relationship");
      const matchingRel = rels.find(r => r.attributes["Id"] === rId);
      expect(matchingRel).toBeDefined();
      expect(matchingRel!.attributes["TargetMode"]).toBe("External");
      expect(matchingRel!.attributes["Target"]).toBe("https://example.com");
    });

    it("should degrade hyperlink without r:id and without anchor to bare runs (avoid CT_Hyperlink schema violation)", async () => {
      const h = Document.create();
      Document.addContent(h, Build.paragraph([Build.hyperlink("plain text", {})]));
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/document.xml");
      const body = findChild(root, "w:body")!;

      // No w:hyperlink should be emitted (since neither r:id nor w:anchor is present).
      const hyps = findDescendants(body, "w:hyperlink");
      expect(hyps.length).toBe(0);

      // The run text must still be present, just unwrapped.
      const xmlText = new TextDecoder().decode(files.get("word/document.xml")!);
      expect(xmlText).toContain("plain text");
    });
  });

  // ===========================================================================
  // 10. 页眉页脚结构合规
  // ===========================================================================

  describe("header/footer structure compliance", () => {
    it("should produce w:hdr/w:ftr root and references in sectPr", async () => {
      const h = Document.create();
      Document.setHeader(h, "default", {
        children: [Build.textParagraph("Header Text")]
      });
      Document.setFooter(h, "default", {
        children: [Build.textParagraph("Footer Text")]
      });
      // Must explicitly set section properties with header/footer refs
      // so that w:headerReference/w:footerReference are emitted in w:sectPr
      Document.setSectionProperties(h, {
        headers: [{ type: "default", rId: "" }],
        footers: [{ type: "default", rId: "" }]
      });
      Document.addParagraph(h, "Body");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      // Find header file
      const headerPath = [...files.keys()].find(
        p => p.startsWith("word/header") && p.endsWith(".xml")
      );
      expect(headerPath).toBeDefined();

      const headerRoot = parseEntry(files, headerPath!);
      expect(headerRoot.name).toBe("w:hdr");

      // Header should contain paragraphs
      const hdrParas = findChildren(headerRoot, "w:p");
      expect(hdrParas.length).toBeGreaterThan(0);

      // Find footer file
      const footerPath = [...files.keys()].find(
        p => p.startsWith("word/footer") && p.endsWith(".xml")
      );
      expect(footerPath).toBeDefined();

      const footerRoot = parseEntry(files, footerPath!);
      expect(footerRoot.name).toBe("w:ftr");

      // Footer should contain paragraphs
      const ftrParas = findChildren(footerRoot, "w:p");
      expect(ftrParas.length).toBeGreaterThan(0);

      // Check sectPr in document.xml has headerReference and footerReference
      const docRoot = parseEntry(files, "word/document.xml");
      const body = findChild(docRoot, "w:body")!;
      const sectPr = findDescendant(body, "w:sectPr");
      expect(sectPr).toBeDefined();

      const headerRef = findChild(sectPr!, "w:headerReference");
      expect(headerRef).toBeDefined();
      expect(headerRef!.attributes["w:type"]).toBe("default");

      const footerRef = findChild(sectPr!, "w:footerReference");
      expect(footerRef).toBeDefined();
      expect(footerRef!.attributes["w:type"]).toBe("default");
    });
  });

  // ===========================================================================
  // 11. 批注结构合规
  // ===========================================================================

  describe("comments structure compliance", () => {
    it("should produce w:comments with w:comment elements and range markers in document", async () => {
      const h = Document.create();
      const commentId = Document.addComment(h, "Author Name", "This is a comment", {
        date: "2024-01-01T00:00:00Z"
      });
      Document.addContent(
        h,
        Build.paragraph([
          Build.commentRangeStart(commentId),
          Build.text("Commented text"),
          Build.commentRangeEnd(commentId),
          Build.commentReference(commentId)
        ])
      );
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      // comments.xml exists
      const commentsPath = "word/comments.xml";
      expect(files.has(commentsPath)).toBe(true);

      const commentsRoot = parseEntry(files, commentsPath);
      expect(commentsRoot.name).toBe("w:comments");

      // Has w:comment elements
      const comments = findChildren(commentsRoot, "w:comment");
      expect(comments.length).toBeGreaterThan(0);

      const comment = comments[0];
      expect(comment.attributes["w:id"]).toBeDefined();
      expect(comment.attributes["w:author"]).toBe("Author Name");
      expect(comment.attributes["w:date"]).toBe("2024-01-01T00:00:00Z");

      // document.xml has commentRangeStart, commentRangeEnd, commentReference
      const docRoot = parseEntry(files, "word/document.xml");
      const rangeStarts = findDescendants(docRoot, "w:commentRangeStart");
      expect(rangeStarts.length).toBeGreaterThan(0);

      const rangeEnds = findDescendants(docRoot, "w:commentRangeEnd");
      expect(rangeEnds.length).toBeGreaterThan(0);

      const refs = findDescendants(docRoot, "w:commentReference");
      expect(refs.length).toBeGreaterThan(0);
    });
  });

  // ===========================================================================
  // 12. settings.xml 结构合规
  // ===========================================================================

  describe("settings.xml structure", () => {
    it("should have w:settings root with namespace", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/settings.xml");

      expect(root.name).toBe("w:settings");

      // Has xmlns:w namespace
      const nsW = root.attributes["xmlns:w"] ?? (root.ns ? root.ns["w"] : undefined);
      expect(nsW).toBeDefined();
    });
  });

  // ===========================================================================
  // 13. fontTable.xml 结构合规
  // ===========================================================================

  describe("fontTable.xml structure", () => {
    it("should have w:fonts root with w:font children having w:name attribute", async () => {
      const h = Document.create();
      Document.useDefaultStyles(h);
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/fontTable.xml");

      expect(root.name).toBe("w:fonts");

      const fonts = findChildren(root, "w:font");
      expect(fonts.length).toBeGreaterThan(0);

      // Each font has w:name attribute
      for (const font of fonts) {
        expect(font.attributes["w:name"]).toBeDefined();
      }
    });
  });

  // ===========================================================================
  // 14. theme/theme1.xml 结构合规
  // ===========================================================================

  describe("theme/theme1.xml structure", () => {
    it("should have a:theme root with a:themeElements containing color, font, and format schemes", async () => {
      const h = Document.create();
      Document.addParagraph(h, "Test");
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      const root = parseEntry(files, "word/theme/theme1.xml");

      expect(root.name).toBe("a:theme");

      const themeElements = findChild(root, "a:themeElements");
      expect(themeElements).toBeDefined();

      const clrScheme = findChild(themeElements!, "a:clrScheme");
      expect(clrScheme).toBeDefined();

      const fontScheme = findChild(themeElements!, "a:fontScheme");
      expect(fontScheme).toBeDefined();

      const fmtScheme = findChild(themeElements!, "a:fmtScheme");
      expect(fmtScheme).toBeDefined();
    });
  });

  // ===========================================================================
  // 15. 图片关系合规
  // ===========================================================================

  describe("image relationship compliance", () => {
    it("should have image Relationship in rels with media file in ZIP and blip reference", async () => {
      const h = Document.create();
      Document.addImage(h, MINI_PNG, "png", 914400, 914400);
      const bytes = await Io.package(Document.build(h));
      const files = await extractDocx(bytes);

      // Find image relationship in document.xml.rels
      const relsRoot = parseEntry(files, "word/_rels/document.xml.rels");
      const rels = findChildren(relsRoot, "Relationship");

      const imageRel = rels.find(r => {
        const type = r.attributes["Type"] ?? "";
        return type.includes("image");
      });
      expect(imageRel).toBeDefined();

      // The target media file exists
      const target = imageRel!.attributes["Target"]!;
      const mediaPath = target.startsWith("/") ? target.substring(1) : `word/${target}`;
      expect(files.has(mediaPath), `Image target "${mediaPath}" should exist in ZIP`).toBe(true);

      // Verify the media file is a valid PNG (starts with PNG signature)
      const mediaData = files.get(mediaPath)!;
      expect(mediaData[0]).toBe(0x89);
      expect(mediaData[1]).toBe(0x50); // P
      expect(mediaData[2]).toBe(0x4e); // N
      expect(mediaData[3]).toBe(0x47); // G

      // document.xml has a:blip with r:embed attribute
      const docRoot = parseEntry(files, "word/document.xml");
      const blips = findDescendants(docRoot, "a:blip");
      expect(blips.length).toBeGreaterThan(0);

      // The blip has an r:embed attribute referencing the image
      const blipWithEmbed = blips.find(b => b.attributes["r:embed"] !== undefined);
      expect(blipWithEmbed).toBeDefined();
    });
  });

  describe("drawing object id (wp:docPr/@id) uniqueness", () => {
    /**
     * Word requires every drawing object id (wp:docPr/@id, and the matching
     * pic:cNvPr/@id) to be a unique positive integer across the *entire*
     * document — body, headers, footers, footnotes, endnotes, comments and
     * text boxes share one id space. Duplicate or zero ids make Word raise the
     * "unreadable content / needs repair" dialog.
     *
     * Regression: floating images (and any inline image whose model carried no
     * `drawingId`) all fell back to a hard-coded id=1, colliding with inline
     * images that did carry ids. See tmp/word-examples/07-images.docx, which
     * emitted eight `<wp:docPr id="1">` elements.
     */
    function collectDocPrIds(files: Map<string, Uint8Array>): { ids: string[]; cNv: string[] } {
      const ids: string[] = [];
      const cNv: string[] = [];
      for (const [path, data] of files) {
        if (!path.endsWith(".xml")) {
          continue;
        }
        if (!/word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(path)) {
          continue;
        }
        const root = parseXml(decoder.decode(data)).root;
        for (const el of findDescendants(root, "wp:docPr")) {
          const id = el.attributes["id"];
          if (id !== undefined) {
            ids.push(id);
          }
        }
        for (const el of findDescendants(root, "pic:cNvPr")) {
          const id = el.attributes["id"];
          if (id !== undefined) {
            cNv.push(id);
          }
        }
      }
      return { ids, cNv };
    }

    it("assigns unique, non-zero ids even when the model has duplicate/unset drawingId", async () => {
      // Two inline images that BOTH claim drawingId=1 (duplicate in the model),
      // plus a floating image with no drawingId at all (would have defaulted to 1).
      const doc = {
        images: [
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image1.png", rId: "rIdImg1" },
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image2.png", rId: "rIdImg2" },
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image3.png", rId: "rIdImg3" }
        ],
        body: [
          Build.paragraph([
            {
              content: [
                {
                  type: "image" as const,
                  rId: "rIdImg1",
                  width: 914400,
                  height: 914400,
                  drawingId: 1
                }
              ]
            }
          ]),
          Build.paragraph([
            {
              content: [
                {
                  type: "image" as const,
                  rId: "rIdImg2",
                  width: 914400,
                  height: 914400,
                  drawingId: 1
                }
              ]
            }
          ]),
          Build.floatingImage({ rId: "rIdImg3", width: 914400, height: 914400 })
        ]
      };

      const bytes = await Io.package(doc);
      const files = await extractDocx(bytes);
      const { ids, cNv } = collectDocPrIds(files);

      expect(ids.length).toBe(3);
      // All unique
      expect(new Set(ids).size).toBe(ids.length);
      // All positive non-zero integers
      for (const id of ids) {
        const n = Number(id);
        expect(Number.isInteger(n)).toBe(true);
        expect(n).toBeGreaterThan(0);
      }
      // pic:cNvPr ids mirror the docPr ids (same logical object)
      expect(new Set(cNv)).toEqual(new Set(ids));
    });

    it("keeps ids unique across body, headers and footers (shared id space)", async () => {
      const headerContent = {
        children: [
          Build.paragraph([
            {
              content: [{ type: "image" as const, rId: "rIdHdr", width: 914400, height: 914400 }]
            }
          ])
        ]
      };
      const footerContent = {
        children: [
          Build.paragraph([
            {
              content: [{ type: "image" as const, rId: "rIdFtr", width: 914400, height: 914400 }]
            }
          ])
        ]
      };

      const doc = {
        images: [
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image1.png", rId: "rIdBody" },
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image2.png", rId: "rIdHdr" },
          { data: MINI_PNG, mediaType: "png" as const, fileName: "image3.png", rId: "rIdFtr" }
        ],
        headers: new Map([["h1", { content: headerContent }]]),
        footers: new Map([["f1", { content: footerContent }]]),
        body: [
          Build.paragraph([
            {
              content: [{ type: "image" as const, rId: "rIdBody", width: 914400, height: 914400 }]
            }
          ])
        ]
      };

      const bytes = await Io.package(doc as never);
      const files = await extractDocx(bytes);
      const { ids } = collectDocPrIds(files);

      // Confirm the header/footer images actually rendered (otherwise the
      // uniqueness check below would pass vacuously).
      const headerRoot = parseEntry(files, "word/header1.xml");
      const footerRoot = parseEntry(files, "word/document.xml");
      void footerRoot;
      expect(findDescendants(headerRoot, "wp:docPr").length).toBe(1);
      expect(findDescendants(parseEntry(files, "word/footer1.xml"), "wp:docPr").length).toBe(1);

      // One image in body, one in a header, one in a footer.
      expect(ids.length).toBe(3);
      expect(new Set(ids).size).toBe(3);
    });
  });
});
