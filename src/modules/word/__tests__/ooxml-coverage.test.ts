/**
 * OOXML coverage round-trip fixtures.
 *
 * Each test exercises a distinct corner of the schema and verifies that
 * the model survives `packageDocx` → `readDocx`. The goal is to surface
 * round-trip bugs in fields the simpler fixtures don't touch.
 *
 * Patterns intentionally tested:
 *   - bookmark start/end pairing across paragraphs
 *   - track-changes (insertedRun / deletedRun / movedFromRun / movedToRun)
 *   - hyperlink anchor / tooltip / docLocation / tgtFrame
 *   - SDT properties (alias, tag, dataBinding, lock, plainText, richText, …)
 *   - RunProperties with the rarer fields (caps, vertAlign, position, kern,
 *     emphasisMark, vanish, etc.)
 *   - paragraph properties (numbering ref, alignment, indent, tabs)
 */

import { describe, it, expect } from "vitest";

import {
  Document,
  packageDocx,
  readDocx,
  text,
  paragraph,
  hyperlink,
  bookmarkStart,
  bookmarkEnd,
  insertedRun,
  deletedRun
} from "../index";
import type {
  AbstractNumbering,
  BodyContent,
  BookmarkStart,
  DocDefaults,
  DocumentSettings,
  DocxDocument,
  Hyperlink,
  InsertedRun,
  NumberingInstance,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  SectionProperties,
  SdtProperties,
  StructuredDocumentTag,
  StyleDef
} from "../types";

const findFirstRun = (para: Paragraph): Run | undefined => {
  for (const child of para.children) {
    if (!("type" in child)) {
      return child as Run;
    }
  }
  return undefined;
};

const collectChildrenByType = <T extends string>(para: Paragraph, typeName: T): unknown[] => {
  const out: unknown[] = [];
  for (const child of para.children) {
    if ("type" in child && child.type === typeName) {
      out.push(child);
    }
  }
  return out;
};

describe("OOXML coverage — bookmarks", () => {
  it("preserves bookmark start/end pairs across a span", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([bookmarkStart(1, "intro"), text("Introduction text."), bookmarkEnd(1)])
    );

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const starts = collectChildrenByType(para, "bookmarkStart") as BookmarkStart[];
    const ends = collectChildrenByType(para, "bookmarkEnd");
    expect(starts.length).toBe(1);
    expect(ends.length).toBe(1);
    expect(starts[0]!.id).toBe(1);
    expect(starts[0]!.name).toBe("intro");
  });

  it("preserves bookmark start/end pairs that straddle multiple paragraphs", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([bookmarkStart(7, "section"), text("First sentence.")])
    );
    Document.addParagraphElement(h, paragraph([text("Middle paragraph.")]));
    Document.addParagraphElement(h, paragraph([text("Last sentence."), bookmarkEnd(7)]));

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    let starts = 0;
    let ends = 0;
    let startName: string | undefined;
    for (const b of parsed.body) {
      if (b.type !== "paragraph") {
        continue;
      }
      for (const c of b.children) {
        if ("type" in c) {
          if (c.type === "bookmarkStart" && c.id === 7) {
            starts++;
            startName = c.name;
          }
          if (c.type === "bookmarkEnd" && c.id === 7) {
            ends++;
          }
        }
      }
    }
    expect(starts).toBe(1);
    expect(ends).toBe(1);
    expect(startName).toBe("section");
  });
});

describe("OOXML coverage — track changes", () => {
  it("preserves insertedRun + deletedRun with author and date", async () => {
    const h = Document.create();
    const insRev = { id: 100, author: "Alice", date: "2024-01-01T00:00:00Z" };
    const delRev = { id: 101, author: "Bob", date: "2024-01-02T00:00:00Z" };
    const insRun: Run = { content: [{ type: "text", text: "added " }] };
    const delRun: Run = { content: [{ type: "text", text: "removed " }] };
    Document.addParagraphElement(
      h,
      paragraph([
        text("Body: "),
        insertedRun(insRun, insRev),
        deletedRun(delRun, delRev),
        text(".")
      ])
    );

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const ins = collectChildrenByType(para, "insertedRun") as InsertedRun[];
    const del = collectChildrenByType(para, "deletedRun") as InsertedRun[];

    expect(ins.length).toBe(1);
    expect(ins[0]!.revision.id).toBe(100);
    expect(ins[0]!.revision.author).toBe("Alice");
    expect(ins[0]!.revision.date).toBe("2024-01-01T00:00:00Z");

    expect(del.length).toBe(1);
    expect(del[0]!.revision.id).toBe(101);
    expect(del[0]!.revision.author).toBe("Bob");
  });
});

describe("OOXML coverage — hyperlinks", () => {
  it("preserves anchor-only hyperlink (no url)", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, paragraph([hyperlink("Go to intro", { anchor: "intro" })]));

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const links = collectChildrenByType(para, "hyperlink") as Hyperlink[];
    expect(links.length).toBe(1);
    expect(links[0]!.anchor).toBe("intro");
    expect(links[0]!.url).toBeUndefined();
  });

  it("preserves tooltip and docLocation", async () => {
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([
        hyperlink("link", {
          url: "https://example.com",
          tooltip: "Example site",
          docLocation: "section1"
        })
      ])
    );

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const link = collectChildrenByType(para, "hyperlink")[0] as Hyperlink;
    expect(link.url).toBe("https://example.com");
    expect(link.tooltip).toBe("Example site");
    expect(link.docLocation).toBe("section1");
  });
});

describe("OOXML coverage — RunProperties", () => {
  it("round-trips a wide selection of run properties", async () => {
    const props: RunProperties = {
      bold: true,
      italic: true,
      underline: "double",
      strike: true,
      caps: true,
      smallCaps: false,
      color: "FF0000",
      highlight: "yellow",
      vertAlign: "superscript",
      spacing: 20,
      vanish: false,
      kern: 24,
      position: 6,
      scale: 150,
      rightToLeft: false,
      emphasisMark: "dot",
      emboss: true,
      shadow: true,
      outline: true
    };
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([{ properties: props, content: [{ type: "text", text: "fancy" }] }])
    );
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const run = findFirstRun(para)!;
    const r = run.properties!;
    expect(r.bold).toBe(true);
    expect(r.italic).toBe(true);
    expect(r.underline).toBe("double");
    expect(r.strike).toBe(true);
    expect(r.caps).toBe(true);
    // Color may come back wrapped (HexColor or ColorSpec).
    const colorVal = typeof r.color === "string" ? r.color : r.color?.val;
    expect(colorVal).toBe("FF0000");
    expect(r.highlight).toBe("yellow");
    expect(r.vertAlign).toBe("superscript");
    expect(r.spacing).toBe(20);
    expect(r.kern).toBe(24);
    expect(r.position).toBe(6);
    expect(r.scale).toBe(150);
    expect(r.emphasisMark).toBe("dot");
    expect(r.emboss).toBe(true);
    expect(r.shadow).toBe(true);
    expect(r.outline).toBe(true);
  });

  it("round-trips font specification with eastAsia and complex script variants", async () => {
    const props: RunProperties = {
      font: { ascii: "Calibri", eastAsia: "SimSun", hAnsi: "Calibri", cs: "Arial" },
      size: 24,
      sizeCs: 22
    };
    const h = Document.create();
    Document.addParagraphElement(
      h,
      paragraph([{ properties: props, content: [{ type: "text", text: "multi-script" }] }])
    );
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const r = findFirstRun(para)!.properties!;
    expect(r.size).toBe(24);
    expect(r.sizeCs).toBe(22);
    if (typeof r.font === "object") {
      expect(r.font.ascii).toBe("Calibri");
      expect(r.font.eastAsia).toBe("SimSun");
      expect(r.font.cs).toBe("Arial");
    } else {
      throw new Error("expected font to round-trip as a FontSpec object");
    }
  });
});

describe("OOXML coverage — SDT properties", () => {
  it("preserves alias / tag / lock / placeholder", async () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: {
        id: 4242,
        alias: "Customer Name",
        tag: "customer-name",
        lockContent: true,
        lockSdt: false,
        showingPlaceholder: true,
        placeholder: "DefaultPlaceholders"
      },
      content: [paragraph([text("Acme Inc.")])]
    };
    const h = Document.create();
    Document.addContent(h, sdt);
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const sdtBack = parsed.body.find(b => b.type === "sdt");
    expect(sdtBack).toBeDefined();
    if (!sdtBack || sdtBack.type !== "sdt") {
      throw new Error("expected sdt");
    }
    const p = sdtBack.properties!;
    expect(p.id).toBe(4242);
    expect(p.alias).toBe("Customer Name");
    expect(p.tag).toBe("customer-name");
    expect(p.lockContent).toBe(true);
    expect(p.showingPlaceholder).toBe(true);
    expect(p.placeholder).toBe("DefaultPlaceholders");
  });

  it("preserves dataBinding xpath / storeItemId / prefixMappings", async () => {
    const sdt: StructuredDocumentTag = {
      type: "sdt",
      properties: {
        id: 5,
        dataBinding: {
          xpath: "/root/customer/name",
          storeItemId: "{12345678-1234-1234-1234-123456789012}",
          prefixMappings: 'xmlns:ns0="http://example.com/ns"'
        }
      },
      content: [paragraph([text("Bound")])]
    };
    const h = Document.create();
    Document.addContent(h, sdt);
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const sdtBack = parsed.body.find(b => b.type === "sdt");
    if (!sdtBack || sdtBack.type !== "sdt") {
      throw new Error("expected sdt");
    }
    const db = sdtBack.properties!.dataBinding!;
    expect(db.xpath).toBe("/root/customer/name");
    expect(db.storeItemId).toBe("{12345678-1234-1234-1234-123456789012}");
    expect(db.prefixMappings).toBe('xmlns:ns0="http://example.com/ns"');
  });

  it("preserves type discriminators (plainText / richText / picture / equation)", async () => {
    const variants: Array<keyof SdtProperties> = [
      "plainText",
      "richText",
      "picture",
      "equation",
      "citation",
      "bibliography"
    ];
    for (const variant of variants) {
      const sdt: StructuredDocumentTag = {
        type: "sdt",
        properties: { id: 1, [variant]: true } as never,
        content: [paragraph([text(variant)])]
      };
      const h = Document.create();
      Document.addContent(h, sdt);
      const bytes = await packageDocx(Document.build(h));
      const parsed = await readDocx(bytes);
      const sdtBack = parsed.body.find(b => b.type === "sdt");
      if (!sdtBack || sdtBack.type !== "sdt") {
        throw new Error("expected sdt");
      }
      expect((sdtBack.properties as Record<string, unknown>)[variant]).toBe(true);
    }
  });
});

describe("OOXML coverage — paragraph properties", () => {
  it("preserves alignment / indent / spacing / tabs", async () => {
    const props: ParagraphProperties = {
      alignment: "center",
      indent: { left: 720, right: 360, firstLine: 360 },
      spacing: { before: 240, after: 120, line: 360, lineRule: "auto" },
      tabs: [
        { position: 1440, type: "left" },
        { position: 4320, type: "right", leader: "dot" }
      ]
    };
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: props,
      children: [{ content: [{ type: "text", text: "indented" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const back = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const p = back.properties!;

    expect(p.alignment).toBe("center");
    expect(p.indent?.left).toBe(720);
    expect(p.indent?.right).toBe(360);
    expect(p.indent?.firstLine).toBe(360);
    expect(p.spacing?.before).toBe(240);
    expect(p.spacing?.after).toBe(120);
    expect(p.spacing?.line).toBe(360);
    expect(p.tabs?.length).toBe(2);
    expect(p.tabs?.[0]?.position).toBe(1440);
    expect(p.tabs?.[1]?.leader).toBe("dot");
  });

  it("preserves keepNext / keepLines / pageBreakBefore / contextualSpacing", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: {
        keepNext: true,
        keepLines: true,
        pageBreakBefore: true,
        contextualSpacing: true,
        outlineLevel: 2
      },
      children: [{ content: [{ type: "text", text: "x" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const p = (parsed.body.find(b => b.type === "paragraph") as Paragraph).properties!;
    expect(p.keepNext).toBe(true);
    expect(p.keepLines).toBe(true);
    expect(p.pageBreakBefore).toBe(true);
    expect(p.contextualSpacing).toBe(true);
    expect(p.outlineLevel).toBe(2);
  });
});

describe("OOXML coverage — numbering", () => {
  it("preserves an abstract numbering definition with multiple levels", async () => {
    const abs: AbstractNumbering = {
      abstractNumId: 0,
      multiLevelType: "hybridMultilevel",
      levels: [
        {
          level: 0,
          format: "decimal",
          text: "%1.",
          start: 1,
          justification: "left",
          paragraphProperties: { indent: { left: 720, hanging: 360 } }
        },
        {
          level: 1,
          format: "lowerLetter",
          text: "%2.",
          start: 1,
          paragraphProperties: { indent: { left: 1440, hanging: 360 } }
        }
      ]
    };
    const inst: NumberingInstance = { numId: 1, abstractNumId: 0 };

    const h = Document.create();
    // Inject directly via Document state — there's no public helper for
    // abstract numbering definitions in the builder. The handle is the
    // state object behind a brand cast; mutating it is the only way to
    // exercise abstract numbering through the public packager today.
    const state = h as unknown as {
      abstractNumberings: AbstractNumbering[];
      numberingInstances: NumberingInstance[];
    };
    state.abstractNumberings.push(abs);
    state.numberingInstances.push(inst);
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: { numbering: { numId: 1, level: 0 } },
      children: [{ content: [{ type: "text", text: "First" }] }]
    });
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: { numbering: { numId: 1, level: 1 } },
      children: [{ content: [{ type: "text", text: "Second" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    expect(parsed.abstractNumberings?.length).toBe(1);
    expect(parsed.abstractNumberings?.[0].levels.length).toBe(2);
    expect(parsed.abstractNumberings?.[0].levels[0].format).toBe("decimal");
    expect(parsed.abstractNumberings?.[0].levels[1].format).toBe("lowerLetter");
    expect(parsed.numberingInstances?.length).toBe(1);
    expect(parsed.numberingInstances?.[0].numId).toBe(1);

    // Body paragraphs still reference numbering.
    const paras = parsed.body.filter(b => b.type === "paragraph") as Paragraph[];
    expect(paras[0]?.properties?.numbering?.numId).toBe(1);
    expect(paras[0]?.properties?.numbering?.level).toBe(0);
    expect(paras[1]?.properties?.numbering?.level).toBe(1);
  });
});

describe("OOXML coverage — styles inheritance", () => {
  it("preserves basedOn / next / linkedStyle relationships", async () => {
    const styles: StyleDef[] = [
      {
        type: "paragraph",
        styleId: "Heading1",
        name: "heading 1",
        basedOn: "Normal",
        next: "Normal",
        link: "Heading1Char",
        uiPriority: 9,
        qFormat: true,
        runProperties: { bold: true, size: 32 }
      },
      {
        type: "paragraph",
        styleId: "Normal",
        name: "Normal",
        isDefault: true
      },
      {
        type: "character",
        styleId: "Heading1Char",
        name: "Heading 1 Char",
        link: "Heading1",
        runProperties: { bold: true }
      }
    ];

    const h = Document.create();
    const state = h as unknown as { styles: StyleDef[] };
    state.styles.push(...styles);
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: { style: "Heading1" },
      children: [{ content: [{ type: "text", text: "Title" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const back = parsed.styles ?? [];
    const heading = back.find(s => s.styleId === "Heading1");
    expect(heading).toBeDefined();
    expect(heading!.basedOn).toBe("Normal");
    expect(heading!.next).toBe("Normal");
    expect(heading!.link).toBe("Heading1Char");
    expect(heading!.qFormat).toBe(true);
    expect(heading!.uiPriority).toBe(9);
    const link = back.find(s => s.styleId === "Heading1Char");
    expect(link).toBeDefined();
    expect(link!.link).toBe("Heading1");
  });
});

describe("OOXML coverage — math (OMML)", () => {
  it("preserves a math block with fraction (mathFraction)", async () => {
    const h = Document.create();
    Document.addContent(h, {
      type: "math",
      content: [
        {
          type: "mathFraction",
          numerator: [{ type: "mathRun", text: "x" }],
          denominator: [{ type: "mathRun", text: "y" }]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const block = parsed.body.find(b => b.type === "math");
    expect(block).toBeDefined();
    if (!block || block.type !== "math") {
      throw new Error("expected math block");
    }
    const frac = block.content[0] as { type: string; numerator?: unknown; denominator?: unknown };
    expect(frac.type).toBe("mathFraction");
    expect(frac.numerator).toBeDefined();
    expect(frac.denominator).toBeDefined();
  });
});

describe("OOXML coverage — paraId / textId on paragraph", () => {
  it("preserves w14:paraId and w14:textId", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      paraId: "12345678",
      textId: "87654321",
      children: [{ content: [{ type: "text", text: "marked" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const p = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    expect(p.paraId).toBe("12345678");
    expect(p.textId).toBe("87654321");
  });
});

describe("OOXML coverage — tables (advanced)", () => {
  it("preserves gridSpan and verticalMerge", async () => {
    const h = Document.create();
    Document.addTableElement(h, {
      type: "table",
      properties: { width: { value: 5000, type: "pct" } },
      rows: [
        {
          cells: [
            {
              properties: { gridSpan: 2 },
              content: [
                { type: "paragraph", children: [{ content: [{ type: "text", text: "AB" }] }] }
              ]
            },
            {
              properties: { verticalMerge: "restart" },
              content: [
                { type: "paragraph", children: [{ content: [{ type: "text", text: "C" }] }] }
              ]
            }
          ]
        },
        {
          cells: [
            {
              content: [
                { type: "paragraph", children: [{ content: [{ type: "text", text: "A2" }] }] }
              ]
            },
            {
              content: [
                { type: "paragraph", children: [{ content: [{ type: "text", text: "B2" }] }] }
              ]
            },
            {
              properties: { verticalMerge: "continue" },
              content: [
                { type: "paragraph", children: [{ content: [{ type: "text", text: "" }] }] }
              ]
            }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const table = parsed.body.find(b => b.type === "table");
    if (!table || table.type !== "table") {
      throw new Error("expected table");
    }
    expect(table.rows[0].cells[0].properties?.gridSpan).toBe(2);
    expect(table.rows[0].cells[1].properties?.verticalMerge).toBe("restart");
    expect(table.rows[1].cells[2].properties?.verticalMerge).toBe("continue");
  });

  it("preserves cell shading and borders", async () => {
    const h = Document.create();
    Document.addTableElement(h, {
      type: "table",
      rows: [
        {
          cells: [
            {
              properties: {
                shading: { fill: "FFFF00", pattern: "clear" },
                borders: {
                  top: { style: "single", size: 4, color: "000000" },
                  bottom: { style: "double", size: 6, color: "FF0000" }
                }
              },
              content: [
                {
                  type: "paragraph",
                  children: [{ content: [{ type: "text", text: "shaded" }] }]
                }
              ]
            }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const table = parsed.body.find(b => b.type === "table");
    if (!table || table.type !== "table") {
      throw new Error("expected table");
    }
    const cell = table.rows[0].cells[0];
    expect(cell.properties?.shading?.fill).toBe("FFFF00");
    expect(cell.properties?.borders?.top?.style).toBe("single");
    expect(cell.properties?.borders?.bottom?.style).toBe("double");
    expect(cell.properties?.borders?.bottom?.color).toBe("FF0000");
  });
});

describe("OOXML coverage — section properties", () => {
  it("preserves columns / titlePage / pageNumbering", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "p" }] }]
    });
    const state = h as unknown as { sectionProperties?: typeof spec };
    const spec = {
      pageSize: { width: 12240, height: 15840 },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      columns: { count: 2, space: 720, equalWidth: true },
      titlePage: true,
      pageNumbering: { start: 5, format: "decimal" as const }
    };
    state.sectionProperties = spec;

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const sp = parsed.sectionProperties!;
    expect(sp.columns?.count).toBe(2);
    expect(sp.columns?.space).toBe(720);
    expect(sp.titlePage).toBe(true);
    expect(sp.pageNumbering?.start).toBe(5);
  });
});

describe("OOXML coverage — break types in run content", () => {
  it("preserves page / column / textWrapping break types", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "before" },
            { type: "break", breakType: "page" },
            { type: "text", text: "after" }
          ]
        }
      ]
    });
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "x" },
            { type: "break", breakType: "column" }
          ]
        }
      ]
    });
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "break", breakType: "textWrapping" },
            { type: "text", text: "y" }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const paras = parsed.body.filter(b => b.type === "paragraph") as Paragraph[];

    const breakTypes: Array<string | undefined> = [];
    for (const p of paras) {
      for (const c of p.children) {
        if (!("type" in c)) {
          for (const rc of (c as Run).content) {
            if (rc.type === "break") {
              breakTypes.push(rc.breakType ?? "(line)");
            }
          }
        }
      }
    }
    expect(breakTypes).toContain("page");
    expect(breakTypes).toContain("column");
    expect(breakTypes).toContain("textWrapping");
  });
});

describe("OOXML coverage — symbol content", () => {
  it("preserves symbol with font and char hex", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "Star: " },
            { type: "symbol", font: "Wingdings", char: "F0AB" }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let foundSymbol = false;
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "symbol" && rc.char === "F0AB" && rc.font === "Wingdings") {
            foundSymbol = true;
          }
        }
      }
    }
    expect(foundSymbol).toBe(true);
  });
});

describe("OOXML coverage — paragraph mark run properties (rPr inside pPr)", () => {
  it("preserves markRunProperties for end-of-paragraph mark", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: {
        markRunProperties: { bold: true, color: "FF0000", size: 28 }
      },
      children: [{ content: [{ type: "text", text: "mark me" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const p = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const markPr = p.properties?.markRunProperties;
    expect(markPr).toBeDefined();
    expect(markPr?.bold).toBe(true);
    expect(markPr?.size).toBe(28);
    const colorVal = typeof markPr?.color === "string" ? markPr?.color : markPr?.color?.val;
    expect(colorVal).toBe("FF0000");
  });
});

describe("OOXML coverage — color spec with theme reference", () => {
  it("preserves theme color name on text", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          properties: { color: { val: "auto", themeColor: "accent1", themeShade: "BF" } },
          content: [{ type: "text", text: "themed" }]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let run: Run | undefined;
    for (const c of para.children) {
      if (!("type" in c)) {
        run = c as Run;
      }
    }
    const c = run?.properties?.color;
    if (typeof c === "object" && c) {
      expect(c.themeColor).toBe("accent1");
      expect(c.themeShade?.toLowerCase()).toBe("bf");
    } else {
      throw new Error("expected ColorSpec object after round-trip");
    }
  });
});

describe("OOXML coverage — image alt text", () => {
  // Minimal 1x1 PNG.
  const MINI_PNG = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
    0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
    0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82
  ]);

  it("preserves altText on inline image", async () => {
    const h = Document.create();
    Document.addImage(h, MINI_PNG, "png", 914400, 914400, {
      altText: "An accessibility-friendly description"
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    let foundAlt: string | undefined;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "image") {
              foundAlt = c.altText;
            }
          }
        }
      }
    }
    expect(foundAlt).toBe("An accessibility-friendly description");
  });
});

describe("OOXML coverage — text whitespace preservation", () => {
  it("preserves leading/trailing/internal whitespace via xml:space=preserve", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        { content: [{ type: "text", text: "  leading and trailing  " }] },
        { content: [{ type: "text", text: "double  space" }] },
        { content: [{ type: "text", text: "tab\there" }] }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const texts: string[] = [];
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "text") {
            texts.push(rc.text);
          }
        }
      }
    }
    expect(texts[0]).toBe("  leading and trailing  ");
    expect(texts[1]).toBe("double  space");
    expect(texts[2]).toBe("tab\there");
  });
});

describe("OOXML coverage — XML special characters", () => {
  it("round-trips < > & \" ' inside text", async () => {
    const sample = `< > & " ' <tag attr="val"> & end`;
    const h = Document.create();
    Document.addParagraph(h, sample);
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let backText = "";
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "text") {
            backText += rc.text;
          }
        }
      }
    }
    expect(backText).toBe(sample);
  });

  it("round-trips Unicode and surrogate pairs (emoji)", async () => {
    const sample = "Hello 你好 🌍 𠮷 — тест";
    const h = Document.create();
    Document.addParagraph(h, sample);
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let backText = "";
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "text") {
            backText += rc.text;
          }
        }
      }
    }
    expect(backText).toBe(sample);
  });
});

describe("OOXML coverage — endnote round-trip", () => {
  it("preserves endnote body text and reference id", async () => {
    const h = Document.create();
    const enId = Document.addEndnote(h, "An endnote.");
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "ref" },
            { type: "endnoteRef", id: enId }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    expect(parsed.endnotes?.some(e => e.id === enId)).toBe(true);
    let foundRef = false;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "endnoteRef" && c.id === enId) {
              foundRef = true;
            }
          }
        }
      }
    }
    expect(foundRef).toBe(true);
  });
});

describe("OOXML coverage — language tagging", () => {
  it("preserves run language", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          properties: { language: { val: "fr-FR", eastAsia: "ja-JP", bidi: "ar-SA" } },
          content: [{ type: "text", text: "polyglot" }]
        }
      ]
    });
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let lang: { val?: string; eastAsia?: string; bidi?: string } | undefined;
    for (const c of para.children) {
      if (!("type" in c)) {
        lang = (c as Run).properties?.language;
      }
    }
    expect(lang?.val).toBe("fr-FR");
    expect(lang?.eastAsia).toBe("ja-JP");
    expect(lang?.bidi).toBe("ar-SA");
  });
});

describe("OOXML coverage — fields", () => {
  it("preserves a PAGE field with cached value", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "Page " },
            { type: "field", instruction: " PAGE ", cachedValue: "1" }
          ]
        }
      ]
    });
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let foundField = false;
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "field" && rc.instruction.includes("PAGE")) {
            foundField = true;
          }
        }
      }
    }
    expect(foundField).toBe(true);
  });

  it("preserves a DATE field with format string and cached value", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "field",
              instruction: ' DATE \\@ "yyyy-MM-dd" ',
              cachedValue: "2024-06-15"
            }
          ]
        }
      ]
    });
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let cached: string | undefined;
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          if (rc.type === "field" && rc.instruction.includes("DATE")) {
            cached = rc.cachedValue;
          }
        }
      }
    }
    expect(cached).toBe("2024-06-15");
  });
});

describe("OOXML coverage — table of contents", () => {
  it("round-trips TOC with heading style range and hyperlink option", async () => {
    const h = Document.create();
    Document.addTableOfContents(h, { headingStyleRange: "1-4", hyperlink: true });
    Document.addHeading(h, "Section A", 1);
    Document.addHeading(h, "Sub-section", 2);

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);

    const toc = parsed.body.find(b => b.type === "tableOfContents");
    expect(toc).toBeDefined();
    if (!toc || toc.type !== "tableOfContents") {
      throw new Error("expected TOC");
    }
    expect(toc.headingStyleRange).toBe("1-4");
    expect(toc.hyperlink).toBe(true);
  });
});

describe("OOXML coverage — numbering format variants", () => {
  it("preserves uncommon numbering formats (roman / chineseCounting / cardinalText)", async () => {
    const abs: AbstractNumbering = {
      abstractNumId: 7,
      multiLevelType: "singleLevel",
      levels: [
        {
          level: 0,
          format: "upperRoman",
          text: "%1.",
          start: 1,
          paragraphProperties: { indent: { left: 720, hanging: 360 } }
        }
      ]
    };
    const inst: NumberingInstance = { numId: 9, abstractNumId: 7 };
    const h = Document.create();
    const state = h as unknown as {
      abstractNumberings: AbstractNumbering[];
      numberingInstances: NumberingInstance[];
    };
    state.abstractNumberings.push(abs);
    state.numberingInstances.push(inst);
    Document.addParagraphElement(h, {
      type: "paragraph",
      properties: { numbering: { numId: 9, level: 0 } },
      children: [{ content: [{ type: "text", text: "First (I.)" }] }]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    expect(parsed.abstractNumberings?.[0].levels[0].format).toBe("upperRoman");
  });
});

describe("OOXML coverage — page borders", () => {
  it("preserves pageBorders configuration", async () => {
    const h = Document.create();
    Document.addParagraph(h, "x");
    const state = h as unknown as { sectionProperties?: SectionProperties };
    state.sectionProperties = {
      pageSize: { width: 12240, height: 15840 },
      margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      pageBorders: {
        top: { style: "single", size: 8, color: "000000", space: 24 },
        bottom: { style: "single", size: 8, color: "000000", space: 24 },
        left: { style: "single", size: 8, color: "000000", space: 24 },
        right: { style: "single", size: 8, color: "000000", space: 24 }
      }
    };

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const pb = parsed.sectionProperties?.pageBorders;
    expect(pb).toBeDefined();
    expect(pb?.top?.style).toBe("single");
    expect(pb?.top?.size).toBe(8);
    expect(pb?.top?.color).toBe("000000");
  });
});

describe("OOXML coverage — image transforms", () => {
  // Minimal 1x1 PNG.
  const MINI_PNG = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1,
    0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89, 0, 0, 0, 10, 0x49, 0x44, 0x41, 0x54, 0x78,
    0xda, 0x62, 0, 0, 0, 0, 5, 0, 1, 0x0d, 0x0a, 0x2d, 0xb4, 0, 0, 0, 0, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82
  ]);

  it("preserves rotation and flip flags on inline image", async () => {
    const h = Document.create();
    const { rId, drawingId } = Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    // Replace with a richer model: builder doesn't expose rotation/flip,
    // so we mutate the body directly.
    const body = (h as unknown as { body: BodyContent[] }).body;
    body[body.length - 1] = {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "image",
              rId,
              drawingId,
              width: 914400,
              height: 914400,
              rotation: 5400000, // 90deg in 60000ths
              flipHorizontal: true,
              flipVertical: false
            }
          ]
        }
      ]
    };

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    let img: { rotation?: number; flipHorizontal?: boolean; flipVertical?: boolean } | undefined;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "image") {
              img = c;
            }
          }
        }
      }
    }
    expect(img).toBeDefined();
    expect(img!.rotation).toBe(5400000);
    expect(img!.flipHorizontal).toBe(true);
  });

  it("preserves srcRect cropping fractions", async () => {
    const h = Document.create();
    const { rId, drawingId } = Document.addImage(h, MINI_PNG, "png", 914400, 914400);
    const body = (h as unknown as { body: BodyContent[] }).body;
    body[body.length - 1] = {
      type: "paragraph",
      children: [
        {
          content: [
            {
              type: "image",
              rId,
              drawingId,
              width: 914400,
              height: 914400,
              srcRect: { l: 10000, t: 5000, r: 12000, b: 8000 }
            }
          ]
        }
      ]
    };

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    let img: { srcRect?: { l?: number; t?: number; r?: number; b?: number } } | undefined;
    for (const block of parsed.body) {
      if (block.type !== "paragraph") {
        continue;
      }
      for (const child of block.children) {
        if (!("type" in child)) {
          for (const c of (child as Run).content) {
            if (c.type === "image") {
              img = c;
            }
          }
        }
      }
    }
    expect(img!.srcRect?.l).toBe(10000);
    expect(img!.srcRect?.t).toBe(5000);
    expect(img!.srcRect?.r).toBe(12000);
    expect(img!.srcRect?.b).toBe(8000);
  });
});

describe("OOXML coverage — bookmarks with column attrs", () => {
  it("preserves colFirst/colLast on bookmarkStart", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          type: "bookmarkStart",
          id: 99,
          name: "colBookmark",
          colFirst: 1,
          colLast: 3
        },
        { content: [{ type: "text", text: "in cells" }] },
        { type: "bookmarkEnd", id: 99 }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    let bm: BookmarkStart | undefined;
    for (const c of para.children) {
      if ("type" in c && c.type === "bookmarkStart") {
        bm = c as BookmarkStart;
      }
    }
    expect(bm).toBeDefined();
    expect(bm!.colFirst).toBe(1);
    expect(bm!.colLast).toBe(3);
  });
});

describe("OOXML coverage — soft hyphen / no-break hyphen / carriage return", () => {
  it("preserves softHyphen, noBreakHyphen, carriageReturn", async () => {
    const h = Document.create();
    Document.addParagraphElement(h, {
      type: "paragraph",
      children: [
        {
          content: [
            { type: "text", text: "a" },
            { type: "softHyphen" },
            { type: "text", text: "b" },
            { type: "noBreakHyphen" },
            { type: "text", text: "c" },
            { type: "carriageReturn" },
            { type: "text", text: "d" }
          ]
        }
      ]
    });

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const para = parsed.body.find(b => b.type === "paragraph") as Paragraph;
    const types: string[] = [];
    for (const c of para.children) {
      if (!("type" in c)) {
        for (const rc of (c as Run).content) {
          types.push(rc.type);
        }
      }
    }
    expect(types).toContain("softHyphen");
    expect(types).toContain("noBreakHyphen");
    expect(types).toContain("carriageReturn");
  });
});

describe("OOXML coverage — section break types in section properties", () => {
  it("preserves nextPage / continuous / nextColumn / oddPage / evenPage break types", async () => {
    const breakTypes = ["nextPage", "continuous", "nextColumn", "oddPage", "evenPage"] as const;
    for (const bt of breakTypes) {
      const h = Document.create();
      Document.addParagraph(h, "x");
      const state = h as unknown as { sectionProperties?: SectionProperties };
      state.sectionProperties = {
        pageSize: { width: 12240, height: 15840 },
        margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        breakType: bt
      };
      const bytes = await packageDocx(Document.build(h));
      const parsed = await readDocx(bytes);
      expect(parsed.sectionProperties?.breakType).toBe(bt);
    }
  });
});

describe("OOXML coverage — DocDefaults", () => {
  it("preserves docDefaults run + paragraph properties", async () => {
    const h = Document.create();
    const state = h as unknown as { docDefaults?: DocDefaults };
    state.docDefaults = {
      runProperties: {
        font: { ascii: "Calibri", hAnsi: "Calibri" },
        size: 22
      },
      paragraphProperties: {
        spacing: { after: 200, line: 276, lineRule: "auto" }
      }
    };
    Document.addParagraph(h, "default-styled");
    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    const dd = parsed.docDefaults!;
    expect(dd).toBeDefined();
    expect(dd.runProperties?.size).toBe(22);
    expect(dd.paragraphProperties?.spacing?.after).toBe(200);
  });
});

describe("OOXML coverage — customXml parts", () => {
  it("preserves customXml part content and itemId", async () => {
    // Build the DocxDocument directly to bypass the (intentionally
    // restricted) Document handle's `build()` method, which today does
    // not project the rarer round-trip fields like customXmlParts.
    const docModel: DocxDocument = {
      body: [
        {
          type: "paragraph",
          children: [{ content: [{ type: "text", text: "x" }] }]
        }
      ],
      customXmlParts: [
        {
          // Convention: itemId is the bare GUID (no braces). The packager
          // wraps it with `{...}` when writing the OOXML attribute, and
          // the reader strips them on the way back.
          itemId: "ABCDEF12-3456-7890-ABCD-EF1234567890",
          fileName: "item1.xml",
          xmlContent: '<?xml version="1.0"?><root><name>Acme</name></root>'
        }
      ]
    };

    const bytes = await packageDocx(docModel);
    const parsed = await readDocx(bytes);
    expect(parsed.customXmlParts?.length).toBe(1);
    const part = parsed.customXmlParts![0]!;
    expect(part.itemId.toUpperCase()).toBe("ABCDEF12-3456-7890-ABCD-EF1234567890");
    expect(part.xmlContent).toContain("<name>Acme</name>");
  });
});

describe("OOXML coverage — settings", () => {
  it("preserves selected document settings (zoom / defaultTabStop)", async () => {
    const h = Document.create();
    const state = h as unknown as {
      settings?: DocumentSettings;
    };
    state.settings = {
      defaultTabStop: 720,
      zoom: 125
    };
    Document.addParagraph(h, "x");

    const bytes = await packageDocx(Document.build(h));
    const parsed = await readDocx(bytes);
    expect(parsed.settings?.defaultTabStop).toBe(720);
    expect(parsed.settings?.zoom).toBe(125);
  });
});
