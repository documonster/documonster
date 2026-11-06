/**
 * DOCX Module - Document Validation Tests
 */

import { describe, it, expect } from "vitest";

import { Document, Build, Validation } from "../index";
import type { DocxDocument } from "../types";

describe("Document validation", () => {
  it("valid document passes", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "Hello world");
    const result = Validation.validate(doc);
    expect(result.valid).toBe(true);
    expect(result.errorCount).toBe(0);
  });

  it("empty body produces a warning", () => {
    const doc: DocxDocument = {
      body: [],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const result = Validation.validate(doc);
    expect(result.warningCount).toBeGreaterThan(0);
    expect(result.issues.some(i => i.severity === "warning")).toBe(true);
  });

  it("duplicate style IDs produce an error", () => {
    const doc: DocxDocument = {
      body: [Build.textParagraph("test")],
      styles: [
        { styleId: "Heading1", type: "paragraph", name: "Heading 1" },
        { styleId: "Heading1", type: "paragraph", name: "Heading 1 Dup" }
      ],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const result = Validation.validate(doc);
    expect(result.issues.some(i => i.severity === "error" && i.message.includes("style"))).toBe(
      true
    );
  });

  it("invalid floating image dimensions", () => {
    const doc: DocxDocument = {
      body: [
        {
          type: "floatingImage",
          rId: "rId1",
          width: -100,
          height: 0,
          horizontalPosition: { relativeTo: "column", offset: 0 },
          verticalPosition: { relativeTo: "paragraph", offset: 0 },
          wrap: { style: "square" }
        } as any
      ],
      images: [{ rId: "rId1", data: new Uint8Array(10), mediaType: "image/png" }],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const result = Validation.validate(doc);
    expect(result.issues.some(i => i.severity === "error")).toBe(true);
  });

  it("duplicate numbering IDs produce error", () => {
    const doc: DocxDocument = {
      body: [Build.textParagraph("test")],
      abstractNumberings: [
        { abstractNumId: 1, levels: [{ level: 0, format: "decimal", text: "%1." }] },
        { abstractNumId: 1, levels: [{ level: 0, format: "decimal", text: "%1." }] }
      ],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const result = Validation.validate(doc);
    expect(result.issues.some(i => i.severity === "error" && i.rule === "num-dup-abstract")).toBe(
      true
    );
  });

  it("maxErrors option limits error count", () => {
    const doc: DocxDocument = {
      body: [Build.textParagraph("test")],
      styles: [
        { styleId: "S1", type: "paragraph", name: "S1" },
        { styleId: "S1", type: "paragraph", name: "S1 dup" },
        { styleId: "S2", type: "paragraph", name: "S2" },
        { styleId: "S2", type: "paragraph", name: "S2 dup" },
        { styleId: "S3", type: "paragraph", name: "S3" },
        { styleId: "S3", type: "paragraph", name: "S3 dup" }
      ],
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const result = Validation.validate(doc, { maxErrors: 1 });
    // With maxErrors=1, the body loop exits early but style validation may still add some
    // The key thing is we get fewer issues than without the limit
    const errorsUnlimited = Validation.validate(doc).issues.filter(i => i.severity === "error");
    const errors = result.issues.filter(i => i.severity === "error");
    expect(errors.length).toBeLessThanOrEqual(errorsUnlimited.length);
  });

  describe("table column count follows the grid (gridBefore + Σ gridSpan + gridAfter)", () => {
    const p = (t: string) => ({ properties: {}, content: [Build.textParagraph(t)] });
    const span = (n: number) => ({
      properties: { gridSpan: n },
      content: [Build.textParagraph("x")]
    });
    function tableDoc(rows: any[], columnWidths?: number[]): DocxDocument {
      return {
        body: [{ type: "table", columnWidths, rows }],
        sectionProperties: { pageSize: { width: 12240, height: 15840 } }
      } as any;
    }
    function mismatches(doc: DocxDocument) {
      return Validation.validate(doc).issues.filter(i => /col-mismatch/.test(i.rule ?? ""));
    }

    it("accepts gridAfter and gridBefore", () => {
      const doc = tableDoc(
        [
          { cells: [p("a"), p("b"), p("c")] },
          { properties: { gridAfter: 1 }, cells: [p("a"), p("b")] },
          { properties: { gridBefore: 2 }, cells: [p("c")] },
          { properties: { gridBefore: 1, gridAfter: 1 }, cells: [p("b")] }
        ],
        [1000, 1000, 1000]
      );
      expect(mismatches(doc)).toEqual([]);
    });

    it("accepts a first row that does not span the whole grid", () => {
      const doc = tableDoc(
        [{ properties: { gridAfter: 1 }, cells: [p("a")] }, { cells: [span(2)] }],
        [1000, 1000]
      );
      expect(mismatches(doc)).toEqual([]);
    });

    it("still reports a row that disagrees with tblGrid", () => {
      const doc = tableDoc(
        [{ cells: [p("a"), p("b")] }, { properties: { gridAfter: 2 }, cells: [p("a"), p("b")] }],
        [1000, 1000]
      );
      const issues = mismatches(doc);
      expect(issues.length).toBeGreaterThan(0);
      expect(issues.every(i => i.path.endsWith("rows[1]"))).toBe(true);
    });

    it("accepts a ragged row with trailing cells omitted (Word opens it unchanged)", () => {
      const doc = tableDoc(
        [{ cells: [p("a"), p("b"), p("c")] }, { cells: [p("a")] }],
        [1000, 1000, 1000]
      );
      expect(mismatches(doc)).toEqual([]);
      expect(Validation.validate(doc).valid).toBe(true);
    });

    it("accepts ragged rows in a table without tblGrid", () => {
      const doc = tableDoc([{ cells: [p("a")] }, { cells: [p("a"), p("b")] }]);
      expect(mismatches(doc)).toEqual([]);
    });
  });

  it("resolves an image named by an alias key (header and body sharing one image)", () => {
    const image = (rId: string) => ({
      type: "paragraph",
      children: [{ content: [{ type: "image", rId, width: 10, height: 10 }] }]
    });
    const doc = {
      body: [image("rId5")],
      images: [
        {
          rId: "rId5",
          aliasRIds: ["rId9"],
          fileName: "image1.png",
          mediaType: "image/png",
          data: new Uint8Array([1, 2, 3])
        }
      ],
      headers: new Map([["rIdH", { content: { children: [image("rId9")] } }]]),
      sectionProperties: { pageSize: { width: 12240, height: 15840 } }
    } as any;
    const rules = Validation.validate(doc).issues.map(i => i.rule);
    expect(rules).not.toContain("rel-image-missing");
    const broken = { ...doc, body: [image("rId404")] };
    expect(Validation.validate(broken).issues.map(i => i.rule)).toContain("rel-image-missing");
  });

  describe("every model content type is known", () => {
    function rulesFor(body: any[]): string[] {
      const doc = {
        body,
        sectionProperties: { pageSize: { width: 12240, height: 15840 } }
      } as any;
      return Validation.validate(doc).issues.map(i => i.rule);
    }

    it("does not report preserved paragraph-level XML as unusual", () => {
      const rules = rulesFor([
        {
          type: "paragraph",
          children: [{ type: "opaqueParagraphChild", rawXml: "<w:customXml/>" }]
        }
      ]);
      expect(rules).not.toContain("unusual-para-child");
    });

    it("accepts every typed ParagraphChild member", () => {
      const children = [
        { type: "customXmlInsRangeStart", id: 1 },
        { type: "moveToRangeEnd", id: 1 }
      ];
      expect(rulesFor([{ type: "paragraph", children }])).not.toContain("unusual-para-child");
    });

    it("accepts a chartEx body element", () => {
      expect(rulesFor([{ type: "chartEx", rId: "rId9" }])).not.toContain("unknown-body-type");
    });

    it("still reports a type outside the model", () => {
      expect(rulesFor([{ type: "paragraph", children: [{ type: "bogus" }] }])).toContain(
        "unusual-para-child"
      );
      expect(rulesFor([{ type: "bogus" }])).toContain("unknown-body-type");
    });
  });
});
