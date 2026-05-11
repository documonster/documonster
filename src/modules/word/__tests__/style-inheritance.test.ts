/**
 * DOCX Module - Full Style Inheritance Tests
 *
 * Tests for resolveRunStyle, resolveNumberingLevel, resolveTableStyle.
 * Existing resolveStyle is covered by other tests.
 */

import { describe, it, expect } from "vitest";

import { resolveRunStyle, resolveNumberingLevel, resolveTableStyle, resolveStyle } from "../index";
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
  } as unknown as DocxDocument;
}

describe("resolveRunStyle", () => {
  it("returns own properties when no style chain", () => {
    const doc = createDoc({});
    const run: Run = {
      properties: { bold: true },
      content: []
    };
    const resolved = resolveRunStyle(doc, run);
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
    const resolved = resolveRunStyle(doc, run);
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

    const resolved = resolveRunStyle(doc, run, { font: "Calibri", size: 22 });
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
    const resolved = resolveRunStyle(doc, run);
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
    const resolved = resolveRunStyle(doc, run);
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
    const resolved = resolveRunStyle(doc, run);
    expect(resolved.chain.length).toBe(2);
  });
});

describe("resolveNumberingLevel", () => {
  it("returns undefined for paragraph without numbering", () => {
    const doc = createDoc({});
    const para: Paragraph = { type: "paragraph", children: [] };
    expect(resolveNumberingLevel(doc, para)).toBeUndefined();
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

    const resolved = resolveNumberingLevel(doc, para);
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
    expect(resolveNumberingLevel(doc, para)).toBeUndefined();
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

    const resolved = resolveNumberingLevel(doc, para);
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
    const resolved = resolveTableStyle(doc, "MyTable");

    expect(resolved.chain).toEqual(["MyTable", "BaseTable"]);
    expect(resolved.runProperties.bold).toBe(true);
    expect(resolved.runProperties.font).toBe("Arial");
  });

  it("returns minimal result for unknown style", () => {
    const doc = createDoc({});
    const resolved = resolveTableStyle(doc, "Unknown");
    expect(resolved.chain).toEqual(["Unknown"]);
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
    const resolved = resolveTableStyle(doc, "T");
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
    const paraStyle = resolveStyle(doc, para);
    expect(paraStyle.runProperties.font).toBe("Calibri");
    expect(paraStyle.runProperties.size).toBe(22);

    // Step 2: resolve run with paragraph context
    const runStyle = resolveRunStyle(doc, run, paraStyle.runProperties);

    // Inherited from paragraph
    expect(runStyle.runProperties.font).toBe("Calibri");
    expect(runStyle.runProperties.size).toBe(22);
    // From run's character style
    expect(runStyle.runProperties.bold).toBe(true);
    // From run's own
    expect(runStyle.runProperties.italic).toBe(true);
  });
});
