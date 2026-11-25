import { extractAll } from "@archive/unzip/extract";
import { Build, Document } from "@word/index";
import type { DocxDocument, StyleDef, Table } from "@word/index";
import { describe, expect, it } from "vitest";

/**
 * What the Open XML SDK's validator — Office's own schema — refuses in files this library wrote, each fixed where it is
 * written. ECMA-376 allows some of it (a table style's row height, both names of one table edge); Office does not,
 * and documents so in [MS-OI29500]. The rest broke ECMA-376 as well.
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function part(bytes: Uint8Array, path: string): Promise<string> {
  const entry = (await extractAll(bytes)).get(path);
  if (!entry) {
    throw new Error(`no ${path}`);
  }
  return new TextDecoder().decode(entry.data);
}

/** A document whose `word/styles.xml` is `styles`, read back. */
async function withStyles(styles: string): Promise<DocxDocument> {
  const base = await Document.toBuffer(Document.create());
  const { Archive } = await import("@archive/index");
  const editor = await Archive.ZipEditor.open(base);
  editor.set(
    "word/styles.xml",
    new TextEncoder().encode(
      `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>${styles}</w:styles>`
    )
  );
  return Document.read(await editor.toBuffer());
}

describe("a table style holds only what Word allows in one", () => {
  const style = (extra: Partial<StyleDef>): StyleDef => ({
    type: "table",
    styleId: "T",
    name: "T",
    ...extra
  });

  it("refuses a field Word does not allow, in the style's table properties and a condition's", async () => {
    for (const bad of [
      style({ tableProperties: { look: { firstRow: true } } }),
      style({ tableProperties: { width: { value: 0, type: "auto" } } }),
      style({ tableStyleConditions: [{ type: "firstRow", tableProperties: { rowBandSize: 1 } }] }),
      style({
        tableStyleConditions: [{ type: "firstRow", rowProperties: { height: { value: 400 } } }]
      }),
      style({ tableStyleConditions: [{ type: "firstRow", cellProperties: { gridSpan: 2 } }] }),
      style({ tableStyleConditions: [{ type: "firstRow", cellProperties: { rowSpan: 2 } }] })
    ]) {
      const doc = Document.create();
      expect(() => Document.addStyle(doc, bad)).toThrow(/as Word reads them/);
      await expect(Document.toBuffer({ body: [], styles: [bad] })).rejects.toThrow(
        /as Word reads them/
      );
    }
  });

  it("writes what Word allows", async () => {
    const good = style({
      tableProperties: { rowBandSize: 1, shading: { fill: "FFFFFF" }, indent: 0 },
      tableStyleConditions: [
        {
          type: "firstRow",
          rowProperties: { cantSplit: true, tableHeader: true },
          cellProperties: { shading: { fill: "000000" }, verticalAlign: "center" }
        }
      ]
    });
    const styles = await part(
      await Document.toBuffer({ body: [], styles: [good] }),
      "word/styles.xml"
    );
    expect(styles).toContain("<w:tblStyleRowBandSize");
    expect(styles).toContain("<w:tblHeader/>");
    expect(styles).toContain('<w:vAlign w:val="center"/>');
  });

  it("sets aside, when reading, what Word does not allow — reported, and not written back", async () => {
    const doc = await withStyles(
      `<w:style w:type="table" w:styleId="T"><w:name w:val="T"/><w:tblPr><w:tblLook w:val="04A0"/><w:tblInd w:w="0" w:type="dxa"/></w:tblPr>` +
        `<w:tblStylePr w:type="firstRow"><w:trPr><w:trHeight w:val="400"/><w:cantSplit/></w:trPr><w:tcPr><w:tcW w:w="10" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr></w:tblStylePr></w:style>`
    );
    const read = doc.styles!.find(s => s.styleId === "T")!;
    const condition = read.tableStyleConditions![0]!;
    expect(read.tableProperties).toMatchObject({ indent: 0 });
    expect(read.tableProperties?.look).toBeUndefined();
    expect(condition.rowProperties).toMatchObject({ cantSplit: true });
    expect(condition.rowProperties?.height).toBeUndefined();
    expect(condition.cellProperties).toMatchObject({ verticalAlign: "center" });
    expect(condition.cellProperties?.width).toBeUndefined();
    const setAside = Document.listUnrecognizedValues(doc).filter(v => !v.writtenBack);
    expect(setAside.map(v => v.value).join("")).toMatch(/w:tblLook[\s\S]*w:trHeight[\s\S]*w:tcW/);

    const styles = await part(await Document.toBuffer(doc), "word/styles.xml");
    const t = styles.slice(styles.indexOf('w:styleId="T"'));
    expect(t).not.toContain("w:tblLook");
    expect(t).not.toContain("w:trHeight");
    expect(t).not.toMatch(/<w:tcPr><w:tcW/);
  });
});

describe("a table edge is named once", () => {
  it("refuses both of an edge's names, in borders and margins", async () => {
    const twips = (value: number) => ({ value, type: "dxa" as const });
    const border = { style: "single" as const, size: 4 };
    for (const properties of [
      { cellMargins: { start: twips(80), left: twips(80) } },
      { borders: { end: border, right: border } }
    ]) {
      expect(() => Build.table([Build.row([Build.cell("x")])], properties)).toThrow(/one edge/);
    }
    expect(() =>
      Build.cell("x", { margins: { start: twips(80), left: twips(80) } } as never)
    ).toThrow(/one edge/);
  });

  it("keeps, when reading, the name a file gives first, and sets the other aside", async () => {
    const base = await Document.toBuffer(Document.create());
    const { Archive } = await import("@archive/index");
    const editor = await Archive.ZipEditor.open(base);
    editor.set(
      "word/document.xml",
      new TextEncoder().encode(
        `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:tbl><w:tblPr><w:tblCellMar><w:left w:w="20" w:type="dxa"/><w:start w:w="10" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`
      )
    );
    const doc = await Document.read(await editor.toBuffer());
    const margins = (doc.body[0] as Table).properties!.cellMargins!;
    expect(margins).toMatchObject({ left: { value: 20, type: "dxa" } });
    expect(margins.start).toBeUndefined();
    expect(Document.listUnrecognizedValues(doc)).toEqual([
      expect.objectContaining({ value: expect.stringContaining("w:start"), writtenBack: false })
    ]);
    const body = await part(await Document.toBuffer(doc), "word/document.xml");
    expect(body).toContain('<w:tblCellMar><w:left w:w="20" w:type="dxa"/></w:tblCellMar>');
  });
});

describe("parts written to the schema", () => {
  it("writes a table's w:tblPr even when it has no properties", async () => {
    const table = Build.table([Build.row([Build.cell("x")])]);
    const { properties: _, ...bare } = table;
    const body = await part(await Document.toBuffer({ body: [bare] }), "word/document.xml");
    expect(body).toMatch(/<w:tbl><w:tblPr\s*\/?>(?:<\/w:tblPr>)?<w:tblGrid>/);
  });

  it("writes a ruby's properties complete, filling in what the model leaves out as Word does", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.paragraph([Build.ruby("漢字", "かんじ")]));
    const body = await part(await Document.toBuffer(doc), "word/document.xml");
    // Word, saving such a ruby in a document whose runs are 12pt, writes these sizes.
    expect(body).toContain(
      '<w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="12"/><w:hpsRaise w:val="22"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr>'
    );
  });

  it("takes a ruby's language from its text when none is given", async () => {
    const lid = async (base: string, ruby: string) => {
      const doc = Document.create();
      Document.addContent(doc, Build.paragraph([Build.ruby(base, ruby)]));
      return /<w:lid w:val="([^"]+)"/.exec(
        await part(await Document.toBuffer(doc), "word/document.xml")
      )![1];
    };
    expect(await lid("漢字", "かんじ")).toBe("ja-JP");
    expect(await lid("한국", "hanguk")).toBe("ko-KR");
    expect(await lid("数据", "shùjù")).toBe("zh-CN");
  });

  it("writes and reads a group character's position as ST_TopBot: bot, not bottom", async () => {
    const doc = Document.create();
    Document.addMath(doc, [
      Build.mathGroupChar([Build.mathRun("x")], { position: "bottom", verticalAlign: "bottom" })
    ]);
    const bytes = await Document.toBuffer(doc);
    const body = await part(bytes, "word/document.xml");
    expect(body).toContain('<m:pos m:val="bot"/>');
    expect(body).toContain('<m:vertJc m:val="bot"/>');
    expect(body).not.toContain('"bottom"');
    expect(JSON.stringify((await Document.read(bytes)).body)).toContain(
      '"position":"bottom","verticalAlign":"bottom"'
    );
  });

  it("refuses a group character's centre alignment, which ST_TopBot does not have", () => {
    expect(() =>
      Build.mathGroupChar([Build.mathRun("x")], { verticalAlign: "center" as never })
    ).toThrow();
  });

  it("writes a delimiter's separator between its begin and end characters", async () => {
    const doc = Document.create();
    Document.addMath(doc, [
      Build.mathDelimiter([[Build.mathRun("a")], [Build.mathRun("b")]], {
        beginChar: "[",
        endChar: "]",
        separatorChar: ";"
      })
    ]);
    const body = await part(await Document.toBuffer(doc), "word/document.xml");
    expect(body).toMatch(/<m:begChr [^>]*\/><m:sepChr [^>]*\/><m:endChr /);
  });

  it("writes a repeating section's title as w:val, the namespace of its type, and reads both forms", async () => {
    const sdt = Build.structuredDocumentTag([Build.textParagraph("Item")], {
      repeatingSection: { sectionTitle: "Item" }
    });
    const bytes = await Document.toBuffer({ body: [sdt] });
    expect(await part(bytes, "word/document.xml")).toContain('<w15:sectionTitle w:val="Item"/>');
    expect(JSON.stringify((await Document.read(bytes)).body)).toContain('"sectionTitle":"Item"');
  });

  it("braces a custom XML part's item id once, given bare or braced", async () => {
    for (const itemId of [
      "{11111111-2222-3333-4444-555555555555}",
      "11111111-2222-3333-4444-555555555555"
    ]) {
      const doc = {
        ...Document.create(),
        customXmlParts: [{ itemId, fileName: "item1.xml", xmlContent: "<a/>" }]
      };
      const entries = await extractAll(await Document.toBuffer(doc));
      const props = [...entries.keys()].find(p => /customXml\/itemProps\d+\.xml$/.test(p))!;
      expect(new TextDecoder().decode(entries.get(props)!.data)).toContain(
        'ds:itemID="{11111111-2222-3333-4444-555555555555}"'
      );
    }
  });

  it("writes legacy compatibility flags in CT_Compat's order", async () => {
    const doc = {
      ...Document.create(),
      settings: {
        compatFlags: [
          { name: "doNotExpandShiftReturn" },
          { name: "useSingleBorderforContiguousCells" }
        ]
      }
    };
    const settings = await part(await Document.toBuffer(doc), "word/settings.xml");
    expect(settings.indexOf("useSingleBorderforContiguousCells")).toBeLessThan(
      settings.indexOf("doNotExpandShiftReturn")
    );
  });
});

describe("chart groups written to their content models", () => {
  async function chartXml(type: string): Promise<string> {
    const bytes = await Document.toBuffer({
      body: [
        {
          type: "chart",
          chart: {
            type,
            series: [{ name: "A", categories: ["1", "2", "3"], values: [1, 2, 3] }]
          }
        } as never
      ]
    });
    const entries = await extractAll(bytes);
    const path = [...entries.keys()].find(p => /word\/charts\/chart\d+\.xml$/.test(p))!;
    return new TextDecoder().decode(entries.get(path)!.data);
  }

  it("writes a cache's number format as its text", async () => {
    const xml = await chartXml("line");
    expect(xml).toContain("<c:formatCode>General</c:formatCode>");
    expect(xml).not.toContain("<c:formatCode val=");
  });

  it("writes varyColors only into a group whose content model has it", async () => {
    expect(await chartXml("surface")).not.toContain("c:varyColors");
    expect(await chartXml("line")).toContain("c:varyColors");
  });

  it("gives a doughnut its first slice angle and hole size", async () => {
    const xml = await chartXml("doughnut");
    expect(xml).toMatch(/<c:firstSliceAng val="0"\/><c:holeSize val="50"\/>/);
  });
});
