/**
 * DOCX Module - Full Style Inheritance Tests
 *
 * Tests for resolveRunStyle, resolveNumberingLevel, resolveTableStyle.
 * Existing resolveStyle is covered by other tests.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, Query } from "../index";
import type {
  AbstractNumbering,
  DocxDocument,
  NumberingInstance,
  Paragraph,
  Run,
  StyleDef
} from "../types";

function createDoc(opts: {
  styles?: StyleDef[];
  abstractNumberings?: AbstractNumbering[];
  numberingInstances?: NumberingInstance[];
  docDefaults?: DocxDocument["docDefaults"];
}): DocxDocument {
  return {
    body: [],
    styles: opts.styles,
    abstractNumberings: opts.abstractNumberings,
    numberingInstances: opts.numberingInstances,
    docDefaults: opts.docDefaults
  };
}

describe("resolveRunStyle", () => {
  it("returns own properties when no style chain", () => {
    const doc = createDoc({});
    const run: Run = {
      properties: { bold: true },
      content: []
    };
    const resolved = Query.getRunStyle(doc, run);
    expect(resolved.runProperties.bold).toBe(true);
    expect(resolved.chain).toEqual([]);
  });

  it("walks character style basedOn chain", () => {
    const styles: StyleDef[] = [
      {
        type: "character",
        styleId: "Strong",
        name: "Strong",
        basedOn: "Default",
        runProperties: { bold: true }
      },
      {
        type: "character",
        styleId: "Default",
        name: "Default Char",
        runProperties: { font: "Arial" }
      }
    ];
    const doc = createDoc({ styles });
    const run: Run = {
      properties: { style: "Strong" },
      content: []
    };
    const resolved = Query.getRunStyle(doc, run);
    expect(resolved.chain).toEqual(["Strong", "Default"]);
    expect(resolved.runProperties.bold).toBe(true);
    expect(resolved.runProperties.font).toBe("Arial");
  });

  it("layers paragraph run properties below character style", () => {
    const styles: StyleDef[] = [
      {
        type: "character",
        styleId: "Bold",
        name: "Bold",
        runProperties: { bold: true }
      }
    ];
    const doc = createDoc({ styles });
    const run: Run = {
      properties: { style: "Bold" },
      content: []
    };

    const resolved = Query.getRunStyle(doc, run, {
      paragraphRunProperties: { font: "Calibri", size: 22 }
    });
    // Inherited from paragraph
    expect(resolved.runProperties.font).toBe("Calibri");
    expect(resolved.runProperties.size).toBe(22);
    // From character style
    expect(resolved.runProperties.bold).toBe(true);
  });

  it("run's own properties take highest priority", () => {
    const styles: StyleDef[] = [
      {
        type: "character",
        styleId: "Red",
        name: "Red",
        runProperties: { color: "FF0000" }
      }
    ];
    const doc = createDoc({ styles });
    const run: Run = {
      properties: { style: "Red", color: "00FF00" },
      content: []
    };
    const resolved = Query.getRunStyle(doc, run);
    // Run's own color overrides style
    expect(resolved.runProperties.color).toBe("00FF00");
  });

  it("merges with doc defaults", () => {
    const doc = createDoc({
      docDefaults: {
        runProperties: { font: "Times" }
      }
    });
    const run: Run = { content: [] };
    const resolved = Query.getRunStyle(doc, run);
    expect(resolved.runProperties.font).toBe("Times");
  });

  it("handles circular basedOn references safely", () => {
    const styles: StyleDef[] = [
      {
        type: "character",
        styleId: "A",
        name: "A",
        basedOn: "B",
        runProperties: { bold: true }
      },
      {
        type: "character",
        styleId: "B",
        name: "B",
        basedOn: "A",
        runProperties: { italic: true }
      }
    ];
    const doc = createDoc({ styles });
    const run: Run = { properties: { style: "A" }, content: [] };
    // Should not infinite loop
    const resolved = Query.getRunStyle(doc, run);
    expect(resolved.chain.length).toBe(2);
  });
});

describe("resolveNumberingLevel", () => {
  it("returns undefined for paragraph without numbering", () => {
    const doc = createDoc({});
    const para: Paragraph = { type: "paragraph", children: [] };
    expect(Query.getNumberingLevel(doc, para)).toBeUndefined();
  });

  it("resolves a simple numbering level", () => {
    const abstractNumberings: AbstractNumbering[] = [
      {
        abstractNumId: 0,
        levels: [
          {
            level: 0,
            format: "decimal",
            text: "%1.",
            justification: "left"
          }
        ]
      }
    ];
    const numberingInstances: NumberingInstance[] = [{ numId: 1, abstractNumId: 0 }];

    const doc = createDoc({ abstractNumberings, numberingInstances });
    const para: Paragraph = {
      type: "paragraph",
      properties: { numbering: { level: 0, numId: 1 } },
      children: []
    };

    const resolved = Query.getNumberingLevel(doc, para);
    expect(resolved).toBeDefined();
    expect(resolved!.format).toBe("decimal");
    expect(resolved!.text).toBe("%1.");
    expect(resolved!.justification).toBe("left");
  });

  it("returns undefined when numbering instance not found", () => {
    const doc = createDoc({ abstractNumberings: [], numberingInstances: [] });
    const para: Paragraph = {
      type: "paragraph",
      properties: { numbering: { level: 0, numId: 999 } },
      children: []
    };
    expect(Query.getNumberingLevel(doc, para)).toBeUndefined();
  });

  it("applies level override", () => {
    const abstractNumberings: AbstractNumbering[] = [
      {
        abstractNumId: 0,
        levels: [{ level: 0, format: "decimal", text: "%1." }]
      }
    ];
    const numberingInstances: NumberingInstance[] = [
      {
        numId: 1,
        abstractNumId: 0,
        overrides: [
          {
            level: 0,
            levelDef: {
              level: 0,
              format: "bullet",
              text: "•"
            }
          }
        ]
      }
    ];

    const doc = createDoc({ abstractNumberings, numberingInstances });
    const para: Paragraph = {
      type: "paragraph",
      properties: { numbering: { level: 0, numId: 1 } },
      children: []
    };

    const resolved = Query.getNumberingLevel(doc, para);
    expect(resolved!.format).toBe("bullet");
    expect(resolved!.text).toBe("•");
  });
});

describe("resolveTableStyle", () => {
  it("walks table style basedOn chain", () => {
    const styles: StyleDef[] = [
      {
        type: "table",
        styleId: "MyTable",
        name: "My Table",
        basedOn: "BaseTable",
        runProperties: { bold: true }
      },
      {
        type: "table",
        styleId: "BaseTable",
        name: "Base Table",
        runProperties: { font: "Arial" }
      }
    ];
    const doc = createDoc({ styles });
    const resolved = Query.getTableStyle(doc, "MyTable");

    expect(resolved.chain).toEqual(["MyTable", "BaseTable"]);
    expect(resolved.runProperties.bold).toBe(true);
    expect(resolved.runProperties.font).toBe("Arial");
  });

  it("returns minimal result for unknown style", () => {
    const doc = createDoc({});
    const resolved = Query.getTableStyle(doc, "Unknown");
    // The chain used to list the unknown id as if it were a style; like a
    // paragraph's, it holds only styles the document defines.
    expect(resolved.chain).toEqual([]);
  });

  it("merges with doc defaults", () => {
    const doc = createDoc({
      docDefaults: {
        runProperties: { font: "Calibri" }
      },
      styles: [
        {
          type: "table",
          styleId: "T",
          name: "T",
          runProperties: { bold: true }
        }
      ]
    });
    const resolved = Query.getTableStyle(doc, "T");
    expect(resolved.runProperties.font).toBe("Calibri");
    expect(resolved.runProperties.bold).toBe(true);
  });
});

// =============================================================================
// Integration: combined paragraph + run resolution
// =============================================================================

describe("combined style resolution", () => {
  it("resolveRunStyle uses resolveStyle's output as base", () => {
    const styles: StyleDef[] = [
      {
        type: "paragraph",
        styleId: "Body",
        name: "Body",
        paragraphProperties: { alignment: "left" },
        runProperties: { font: "Calibri", size: 22 }
      },
      {
        type: "character",
        styleId: "Strong",
        name: "Strong",
        runProperties: { bold: true }
      }
    ];
    const doc = createDoc({ styles });

    const para: Paragraph = {
      type: "paragraph",
      properties: { style: "Body" },
      children: []
    };
    const run: Run = {
      properties: { style: "Strong", italic: true },
      content: []
    };

    // Step 1: resolve paragraph style
    const paraStyle = Query.getParagraphStyle(doc, para);
    expect(paraStyle.runProperties.font).toBe("Calibri");
    expect(paraStyle.runProperties.size).toBe(22);

    // Step 2: resolve run with paragraph context
    const runStyle = Query.getRunStyle(doc, run, {
      paragraphRunProperties: paraStyle.runProperties
    });

    // Inherited from paragraph
    expect(runStyle.runProperties.font).toBe("Calibri");
    expect(runStyle.runProperties.size).toBe(22);
    // From run's character style
    expect(runStyle.runProperties.bold).toBe(true);
    // From run's own
    expect(runStyle.runProperties.italic).toBe(true);
  });
});

describe("attribute-level inheritance of property bags", () => {
  it("keeps the document default line spacing when a style sets only `after`", () => {
    // `w:spacing` is a bag of independently inherited attributes. Replacing the
    // whole object made `ListParagraph` — which declares only `after` — drop the
    // document default's `line`, so every Markdown list item rendered at single
    // spacing (13.2pt) while the paragraphs around it were at 1.31 (17.27pt).
    const doc = createDoc({
      docDefaults: {
        paragraphProperties: { spacing: { after: 251, line: 314, lineRule: "auto" } }
      },
      styles: [
        { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true },
        {
          type: "paragraph",
          styleId: "ListParagraph",
          name: "List Paragraph",
          basedOn: "Normal",
          paragraphProperties: { contextualSpacing: true, spacing: { after: 154 } }
        }
      ]
    });
    const para: Paragraph = {
      type: "paragraph",
      properties: { style: "ListParagraph" },
      children: []
    };

    const { spacing } = Query.getParagraphStyle(doc, para).paragraphProperties;
    expect(spacing).toEqual({ after: 154, line: 314, lineRule: "auto" });
  });

  it("merges indentation, borders and shading per attribute", () => {
    const doc = createDoc({
      styles: [
        {
          type: "paragraph",
          styleId: "Base",
          name: "Base",
          paragraphProperties: {
            indent: { left: 720, firstLine: 360, right: 240 },
            borders: {
              top: { style: "single", size: 4, color: "AAAAAA" },
              bottom: { style: "single", size: 4 }
            },
            shading: { fill: "FFFFFF", pattern: "clear", color: "auto" }
          }
        },
        {
          type: "paragraph",
          styleId: "Derived",
          name: "Derived",
          basedOn: "Base",
          // Only one member of each bag — the rest must survive.
          paragraphProperties: {
            indent: { left: 1440 },
            borders: { bottom: { style: "double" } },
            shading: { fill: "EEEEEE" }
          }
        }
      ]
    });
    const para: Paragraph = { type: "paragraph", properties: { style: "Derived" }, children: [] };

    const props = Query.getParagraphStyle(doc, para).paragraphProperties;
    expect(props.indent).toEqual({ left: 1440, firstLine: 360, right: 240 });
    // The container merges per side; a side is still replaced whole, because a
    // `Border` requires `style` and so can never arrive half-declared.
    expect(props.borders).toEqual({
      top: { style: "single", size: 4, color: "AAAAAA" },
      bottom: { style: "double" }
    });
    expect(props.shading).toEqual({ fill: "EEEEEE", pattern: "clear", color: "auto" });
  });

  it("merges w:rFonts per script slot but replaces a bare family name", () => {
    const doc = createDoc({
      docDefaults: { runProperties: { font: { ascii: "Calibri", eastAsia: "MS Mincho" } } },
      styles: [
        {
          type: "character",
          styleId: "Latin",
          name: "Latin",
          runProperties: { font: { ascii: "Arial", hAnsi: "Arial" } }
        },
        {
          type: "character",
          styleId: "Named",
          name: "Named",
          runProperties: { font: "Courier New" }
        }
      ]
    });

    const merged = Query.getRunStyle(doc, { properties: { style: "Latin" }, content: [] });
    expect(merged.runProperties.font).toEqual({
      ascii: "Arial",
      hAnsi: "Arial",
      eastAsia: "MS Mincho"
    });

    // A `FontSpec` and a family name are different shapes, so the name wins whole.
    const named = Query.getRunStyle(doc, { properties: { style: "Named" }, content: [] });
    expect(named.runProperties.font).toBe("Courier New");
  });

  it("does not merge scalars that share a key name with a property bag", () => {
    // `RunProperties.spacing` is character spacing in twips, not a `w:spacing`
    // bag — a single key set is only safe because a scalar never takes the
    // merge branch.
    const doc = createDoc({
      docDefaults: { runProperties: { spacing: 20 } },
      styles: [
        {
          type: "character",
          styleId: "Tight",
          name: "Tight",
          runProperties: { spacing: -10 }
        }
      ]
    });
    const resolved = Query.getRunStyle(doc, { properties: { style: "Tight" }, content: [] });
    expect(resolved.runProperties.spacing).toBe(-10);
  });

  it("treats an absent member as inherit, not reset", () => {
    const doc = createDoc({
      docDefaults: { paragraphProperties: { spacing: { before: 100, after: 200 } } },
      styles: [
        {
          type: "paragraph",
          styleId: "Explicit",
          name: "Explicit",
          // An explicit 0 still wins — only `undefined` defers.
          paragraphProperties: { spacing: { after: 0, before: undefined } }
        }
      ]
    });
    const para: Paragraph = { type: "paragraph", properties: { style: "Explicit" }, children: [] };
    expect(Query.getParagraphStyle(doc, para).paragraphProperties.spacing).toEqual({
      before: 100,
      after: 0
    });
  });
});

describe("default styles", () => {
  const styles: StyleDef[] = [
    { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true },
    {
      type: "paragraph",
      styleId: "Body",
      name: "Body",
      isDefault: true,
      paragraphProperties: { spacing: { after: 300 } },
      runProperties: { size: 30 }
    },
    {
      type: "character",
      styleId: "Plain",
      name: "Plain",
      isDefault: true,
      runProperties: { color: { val: "FF0000" } }
    }
  ];

  it("gives a paragraph naming no style the last default paragraph style, as Word does", () => {
    const para: Paragraph = { type: "paragraph", children: [] };
    const resolved = Query.getParagraphStyle(createDoc({ styles }), para);
    expect(resolved.chain).toEqual(["Body"]);
    expect(resolved.paragraphProperties.spacing).toEqual({ after: 300 });
    expect(resolved.runProperties.size).toBe(30);
  });

  it("gives a paragraph naming an undefined style the default paragraph style", () => {
    const para: Paragraph = { type: "paragraph", properties: { style: "Missing" }, children: [] };
    expect(Query.getParagraphStyle(createDoc({ styles }), para).chain).toEqual(["Body"]);
  });

  it("gives a run naming no character style the default character style", () => {
    const run: Run = { content: [] };
    const resolved = Query.getRunStyle(createDoc({ styles }), run);
    expect(resolved.chain).toEqual(["Plain"]);
    expect(resolved.runProperties.color).toEqual({ val: "FF0000" });
  });
});

describe("a style id the document does not define", () => {
  it("is not reported in a style chain", () => {
    const doc = Document.create();
    Document.addStyle(doc, {
      styleId: "Child",
      type: "paragraph",
      name: "Child",
      basedOn: "Missing"
    });
    Document.addStyle(doc, {
      styleId: "TChild",
      type: "table",
      name: "TChild",
      basedOn: "Missing"
    });
    expect(Query.getTableStyle(doc, "Nope").chain).toEqual([]);
    expect(Query.getTableStyle(doc, "TChild").chain).toEqual(["TChild"]);
    const para = Build.textParagraph("x", { style: "Child" });
    expect(Query.getParagraphStyle(doc, para).chain).toEqual(["Child"]);
  });
});
