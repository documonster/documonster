import { extractAll } from "@archive/unzip/extract";
import { ZipArchive } from "@archive/zip/zip-archive";
import { Build, Document, DocxArgumentError, DocxWriteError, Validation } from "@word/index";
import type { DocxDocument, Paragraph, StyleDef, Table } from "@word/index";
import {
  parseTableCellProperties,
  parseTableProperties,
  parseTableRowProperties
} from "@word/reader/table-properties-parsers";
import { OOXML_SEQUENCES } from "@word/writer/ooxml-sequences";
import { checkDocxElementOrder, DOCX_ORDER_PARTS } from "@word/writer/order-check";
import { SETTINGS_ELEMENTS } from "@word/writer/parts-writer";
import { emitInSchemaOrder } from "@word/writer/schema-order";
import { parseXml } from "@xml/dom";
import { CHART_ORDER_RULES, CHART_SERIES_ORDER_RULES } from "@xml/ooxml-chart-order";
import type { SequenceNode } from "@xml/sequence-order";
import { describe, expect, it } from "vitest";

/**
 * The writer emits each container's children in the order its ECMA-376 content model requires. A one-line `if`
 * placed too low used to decide that order, and `w:pPr`, `w:tblPr`, the borders and margins, a table style's
 * `w:tblPr` and `w:tcPr`, and `w:abstractNum` were each written out of order somewhere. This document carries
 * every one of those at once.
 */

async function partsOf(bytes: Uint8Array): Promise<Map<string, string>> {
  const parts = new Map<string, string>();
  for (const [path, entry] of await extractAll(bytes)) {
    if (path.endsWith(".xml")) {
      parts.set(path, new TextDecoder().decode(entry.data));
    }
  }
  return parts;
}

const border = { style: "single", size: 4 } as const;
const twips = (value: number) => ({ value, type: "dxa" }) as const;

const paragraph: Paragraph = Build.paragraph([Build.text("Every property")], {
  style: "Normal",
  keepNext: true,
  numbering: { numId: 1, level: 0 },
  borders: { bottom: border },
  shading: { fill: "EEEEEE" },
  tabs: [{ type: "left", position: 720 }],
  suppressAutoHyphens: true,
  kinsoku: false,
  wordWrap: false,
  autoSpaceEastAsianText: false,
  bidi: true,
  snapToGrid: false,
  spacing: { before: 120, after: 120 },
  indent: { left: 720 },
  contextualSpacing: true,
  alignment: "center",
  textAlignment: "center",
  outlineLevel: 1,
  cnfStyle: "000000000000",
  markRunProperties: { bold: true },
  propertyChange: { revision: { id: 1, author: "A" }, previousProperties: { alignment: "left" } }
});

const table: Table = Build.table(
  [
    Build.row([
      Build.cell("Cell", {
        borders: { top: border, start: border, insideH: border, tl2br: border },
        margins: { top: twips(40), start: twips(80), end: twips(80) },
        shading: { fill: "DDDDDD" }
      })
    ])
  ],
  {
    style: "Grid",
    visuallyRightToLeft: true,
    rowBandSize: 1,
    width: { value: 5000, type: "pct" },
    borders: { top: border, start: border, insideH: border, insideV: border, end: border },
    cellMargins: { top: twips(40), start: twips(80), end: twips(80) },
    look: { firstRow: true },
    columnWidths: [9000]
  }
);

const tableStyle: StyleDef = {
  type: "table",
  styleId: "Grid",
  name: "Grid",
  // What Word allows in a style's table properties (`TABLE_STYLE_EXCLUDED`), each child of `CT_TblPrBase` it keeps.
  tableProperties: {
    rowBandSize: 1,
    indent: 0,
    borders: { top: border, start: border, insideH: border },
    cellMargins: { top: twips(0), start: twips(108) },
    cellSpacing: twips(0),
    shading: { fill: "FFFFFF" }
  },
  tableStyleConditions: [
    {
      type: "firstRow",
      cellProperties: { shading: { fill: "000000" }, borders: { bottom: border } }
    }
  ]
};

const control = Build.structuredDocumentTag([Build.textParagraph("Name")], {
  id: 7,
  tag: "name",
  alias: "Customer Name",
  lockContent: true,
  appearance: "boundingBox",
  placeholder: "DefaultPlaceholder",
  temporary: true,
  showingPlaceholder: true,
  dataBinding: { xpath: "/root/name", storeItemId: "{00000000-0000-0000-0000-000000000000}" },
  plainText: true
});

/** The two Word extensions inside `w:sdtPr`: a checkbox and a repeating section, each with every child. */
const extensions = [
  Build.structuredDocumentTag([Build.textParagraph("☐")], {
    tag: "agree",
    checkbox: {
      checked: true,
      checkedChar: "2612",
      checkedFont: "MS Gothic",
      uncheckedChar: "2610",
      uncheckedFont: "MS Gothic"
    }
  }),
  Build.structuredDocumentTag([Build.textParagraph("Item")], {
    tag: "items",
    repeatingSection: { sectionTitle: "Item", allowInsertDelete: false }
  })
];

const document: DocxDocument = {
  body: [paragraph, table, control, ...extensions],
  styles: [tableStyle],
  abstractNumberings: [
    {
      abstractNumId: 1,
      multiLevelType: "multilevel",
      numStyleLink: "ListB",
      styleLink: "ListA",
      levels: [{ level: 0, format: "decimal", text: "%1." }]
    }
  ],
  numberingInstances: [{ numId: 1, abstractNumId: 1 }]
};

describe("DOCX child element order", () => {
  it("writes every container the writer emits in its content model's order", async () => {
    const parts = await partsOf(await Document.toBuffer(document));
    expect(checkDocxElementOrder(parts)).toEqual([]);

    // And the positions that used to be wrong, by name, so a broken order check cannot pass this vacuously.
    const body = parts.get("word/document.xml")!;
    const at = (xml: string, tag: string) => xml.indexOf(`<${tag}`);
    expect(at(body, "w:numPr")).toBeLessThan(at(body, "w:spacing"));
    expect(at(body, "w:kinsoku")).toBeLessThan(at(body, "w:spacing"));
    expect(at(body, "w:cnfStyle")).toBeLessThan(body.indexOf("<w:rPr><w:b/>"));
    expect(at(body, "w:bidiVisual")).toBeLessThan(at(body, "w:tblStyleRowBandSize"));
    expect(at(body, "w:alias")).toBeLessThan(at(body, "w:id "));
    expect(at(body, "w:placeholder")).toBeLessThan(at(body, "w:showingPlcHdr"));
    const numbering = parts.get("word/numbering.xml")!;
    expect(at(numbering, "w:styleLink")).toBeLessThan(at(numbering, "w:numStyleLink"));
  });

  it("keeps the start and end margins and borders of a table style", async () => {
    const styles = (await partsOf(await Document.toBuffer(document))).get("word/styles.xml")!;
    const tblPr = styles.slice(styles.indexOf("<w:tblPr>"), styles.indexOf("</w:tblPr>"));
    expect(tblPr).toContain("<w:start ");
    expect(tblPr).toMatch(/<w:tblCellMar><w:top [^>]*\/><w:start /);
  });

  it("refuses a diagonal on a table's borders, which only a cell has", async () => {
    const diagonal = Build.table([Build.row([Build.cell("x")])], { columnWidths: [100] });
    const withDiagonal: Table = {
      ...diagonal,
      properties: { ...diagonal.properties, borders: { tl2br: border } }
    };
    await expect(Document.toBuffer({ body: [withDiagonal] })).rejects.toThrow(DocxArgumentError);
  });

  it("sets aside a diagonal found in a table's borders when reading, rather than typing it", () => {
    const tblPr = parseXml(
      `<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblBorders><w:top w:val="single"/><w:tl2br w:val="single"/></w:tblBorders></w:tblPr>`
    ).root;
    const borders = parseTableProperties(tblPr).borders!;
    expect(borders.top).toBeDefined();
    expect(borders.tl2br).toBeUndefined();
  });

  it("reads, reports and saves a file whose table borders carry a diagonal", async () => {
    // What an earlier version of this library wrote when a caller set one: `w:tl2br` inside `w:tblBorders`.
    const plain = await Document.toBuffer({
      body: [
        Build.table([Build.row([Build.cell("x")])], {
          borders: { top: border },
          columnWidths: [100]
        })
      ]
    });
    const files = new Map<string, Uint8Array>();
    for (const [path, entry] of await extractAll(plain)) {
      files.set(path, entry.data);
    }
    const body = new TextDecoder().decode(files.get("word/document.xml")!);
    const injected = body.replace("</w:tblBorders>", '<w:tl2br w:val="single"/></w:tblBorders>');
    expect(injected).not.toBe(body);
    files.set("word/document.xml", new TextEncoder().encode(injected));
    const archive = new ZipArchive();
    for (const [path, data] of files) {
      archive.add(path, data);
    }

    const read = await Document.read(await archive.toBuffer());
    const kept = Document.listUnrecognizedValues(read).find(entry =>
      entry.path.endsWith("borders.tl2br")
    );
    expect(kept).toMatchObject({ writtenBack: false, value: expect.stringContaining("tl2br") });
    const issue = Validation.validate(read).issues.find(i => i.path === kept!.path);
    expect(issue?.message).toContain("is not written back");
    const written = await partsOf(await Document.toBuffer(read));
    expect(written.get("word/document.xml")).not.toContain("tl2br");
    expect(checkDocxElementOrder(written)).toEqual([]);
  });

  it("reports a paragraph property written out of order", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:spacing/><w:numPr/></w:pPr></w:p></w:body></w:document>`;
    expect(checkDocxElementOrder(new Map([["word/document.xml", xml]]))).toEqual([
      {
        part: "word/document.xml",
        message: expect.stringContaining(
          "<document/body/p/pPr> child <numPr> appears after <spacing>"
        )
      }
    ]);
  });

  it("places every generated content model somewhere", () => {
    const used = new Set<string>();
    const seen = new Set<SequenceNode<string>>();
    const visit = (node: SequenceNode<string>) => {
      if (seen.has(node)) {
        return;
      }
      seen.add(node);
      if (node.type) {
        used.add(node.type);
      }
      Object.values(node.children ?? {}).forEach(visit);
    };
    Object.values(DOCX_ORDER_PARTS).forEach(visit);
    // Chart content models are placed by the rules shared with the workbook validator.
    for (const rule of [...CHART_ORDER_RULES, ...CHART_SERIES_ORDER_RULES]) {
      used.add(rule.type);
    }
    expect(Object.keys(OOXML_SEQUENCES).filter(type => !used.has(type))).toEqual([]);
  });
});

describe("the preserved-element order of w:settings", () => {
  it("follows CT_Settings, apart from the one element it keeps outside the schema", () => {
    const written = SETTINGS_ELEMENTS.filter(name => name !== "w:doNotDemoteNonCombiningChars");
    // The table names a child of another namespace with its prefix (`m:mathPr`), and its own children without.
    const schema = OOXML_SEQUENCES["w:CT_Settings"]
      .flat()
      .map((name: string) => (name.includes(":") ? name : `w:${name}`));
    // Every name the writer lists is one the schema has, in its namespace and in the schema's relative order.
    expect(written.filter(name => !schema.includes(name))).toEqual([]);
    expect(written).toEqual(schema.filter(name => written.includes(name)));
  });
});

describe("chart rules shared with the workbook validator", () => {
  it("name only content models this module's table has", () => {
    const keys = new Set<string>(Object.keys(OOXML_SEQUENCES));
    expect(
      [...CHART_ORDER_RULES, ...CHART_SERIES_ORDER_RULES]
        .map(rule => rule.type)
        .filter(type => !keys.has(type))
    ).toEqual([]);
  });

  it("report a trendline whose children are out of order", () => {
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:ser><c:trendline><c:trendlineType val="linear"/><c:dispEq val="1"/><c:dispRSqr val="1"/></c:trendline></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;
    expect(checkDocxElementOrder(new Map([["word/charts/chart1.xml", xml]]))).toEqual([
      {
        part: "word/charts/chart1.xml",
        message: expect.stringContaining("child <dispRSqr> appears after <dispEq>")
      }
    ]);
  });
});

describe("emitInSchemaOrder", () => {
  const emitted = (children: Parameters<typeof emitInSchemaOrder>[1]) => {
    const names: string[] = [];
    emitInSchemaOrder(
      "w:CT_SdtPr",
      children.map(child => child && ([child[0], () => names.push(child[0])] as const))
    );
    return names;
  };

  it("sorts by the content model, whatever order the writer lists", () => {
    expect(
      emitted([
        ["w:id", () => {}],
        ["w:tag", () => {}],
        ["w:alias", () => {}]
      ])
    ).toEqual(["w:alias", "w:tag", "w:id"]);
  });

  it("keeps an extension element beside the standard child before it", () => {
    expect(
      emitted([
        ["w:showingPlcHdr", () => {}],
        ["w:lock", () => {}],
        ["w15:appearance", () => {}],
        ["w:alias", () => {}]
      ])
    ).toEqual(["w:alias", "w:lock", "w15:appearance", "w:showingPlcHdr"]);
  });

  it("refuses a name the content model does not have, or one from another namespace", () => {
    expect(() => emitted([["w:bogus", () => {}]])).toThrow(DocxWriteError);
    expect(() => emitted([["a:alias", () => {}]])).toThrow(DocxWriteError);
  });
});

describe("a diagonal on a table's borders", () => {
  const diagonal = { tl2br: border };

  it("is refused in a row's table-property exceptions", async () => {
    const plain = Build.table([Build.row([Build.cell("x")])], { columnWidths: [100] });
    const withException: Table = {
      ...plain,
      rows: [{ ...plain.rows[0]!, properties: { tblPrEx: { borders: diagonal } } }]
    };
    await expect(Document.toBuffer({ body: [withException] })).rejects.toThrow(DocxArgumentError);
  });

  it("is refused in a table style", async () => {
    const style: StyleDef = {
      type: "table",
      styleId: "D",
      name: "D",
      tableProperties: { borders: diagonal }
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
  });
});

describe("a table style's properties", () => {
  /** Read back what was written, so the comparison is model to model. */
  async function roundTrip(style: StyleDef): Promise<StyleDef> {
    const read = await Document.read(await Document.toBuffer({ body: [], styles: [style] }));
    return read.styles!.find(s => s.styleId === style.styleId)!;
  }

  const themed = {
    style: "single",
    size: 8,
    color: "4472C4",
    themeColor: "accent1",
    shadow: true
  } as const;

  it("keeps everything a border carries, and invents nothing it does not", async () => {
    const back = await roundTrip({
      type: "table",
      styleId: "Themed",
      name: "Themed",
      tableProperties: { borders: { top: themed, bottom: { style: "single" } } }
    });
    expect(back.tableProperties?.borders?.top).toMatchObject(themed);
    // No size was given, so none is written — a style border used to come back as `sz="4"`, `color="auto"`.
    expect(back.tableProperties?.borders?.bottom?.size).toBeUndefined();
    expect(back.tableProperties?.borders?.bottom?.color).toBeUndefined();
  });

  it("writes a conditional format's table, row and cell properties as a table's", async () => {
    const style: StyleDef = {
      type: "table",
      styleId: "Banded",
      name: "Banded",
      tableStyleConditions: [
        {
          type: "firstRow",
          tableProperties: { shading: { fill: "D9E2F3" }, cellMargins: { top: twips(40) } },
          rowProperties: { cantSplit: true, tableHeader: true },
          cellProperties: {
            borders: { bottom: themed },
            shading: { fill: "1F3864" },
            verticalAlign: "center",
            margins: { start: twips(60) }
          }
        }
      ]
    };
    const back = await roundTrip(style);
    const condition = back.tableStyleConditions![0]!;
    expect(condition.tableProperties?.shading?.fill).toBe("D9E2F3");
    expect(condition.tableProperties?.cellMargins?.top).toEqual(twips(40));
    expect(condition.rowProperties).toEqual({ cantSplit: true, tableHeader: true });
    expect(condition.cellProperties).toMatchObject({
      borders: { bottom: themed },
      verticalAlign: "center",
      margins: { start: twips(60) }
    });

    // A style's cell properties get no `w:tcW` placeholder: that would make every cell it formats auto-width.
    const styles = (await partsOf(await Document.toBuffer({ body: [], styles: [style] }))).get(
      "word/styles.xml"
    )!;
    const tcPr = styles.slice(styles.indexOf("<w:tblStylePr"), styles.indexOf("</w:tblStylePr>"));
    expect(tcPr).toContain("<w:tcPr>");
    expect(tcPr).not.toContain("w:tcW");
  });

  it("refuses a change record on a style's table properties, which CT_TblPrBase does not have", async () => {
    const style: StyleDef = {
      type: "table",
      styleId: "Changed",
      name: "Changed",
      tableProperties: { propertyChange: { revision: { id: 1, author: "A" } } }
    };
    await expect(Document.toBuffer({ body: [], styles: [style] })).rejects.toThrow(
      DocxArgumentError
    );
  });
});

describe("Word extension order", () => {
  it("reports a w14:checkbox whose state precedes its value", () => {
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:sdt><w:sdtPr><w14:checkbox><w14:checkedState w14:val="2612"/><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr></w:sdt></w:body></w:document>`;
    expect(checkDocxElementOrder(new Map([["word/document.xml", xml]]))).toEqual([
      {
        part: "word/document.xml",
        message: expect.stringContaining("child <checked> appears after <checkedState>")
      }
    ]);
  });
});

describe("emitInSchemaOrder's extension namespaces", () => {
  const names = (children: Parameters<typeof emitInSchemaOrder>[1]) => {
    const out: string[] = [];
    emitInSchemaOrder(
      "w:CT_SdtPr",
      children.map(child => child && ([child[0], () => out.push(child[0])] as const))
    );
    return out;
  };

  it("treats the w16 family as extensions", () => {
    expect(
      names([
        ["w:tag", () => {}],
        ["w16cid:durableId", () => {}],
        ["w:alias", () => {}]
      ])
    ).toEqual(["w:alias", "w:tag", "w16cid:durableId"]);
  });

  it("does not take w10: (VML) for one", () => {
    expect(() => names([["w10:wrap", () => {}]])).toThrow(DocxWriteError);
  });
});

describe("a table style's shading", () => {
  const shading = { fill: "D9E2F3" } as const;
  const style: StyleDef = {
    type: "table",
    styleId: "Shaded",
    name: "Shaded",
    tableProperties: { shading },
    tableStyleConditions: [{ type: "firstRow", cellProperties: { shading } }]
  };

  it("is written as a table's: pattern clear and colour auto unless given", async () => {
    const parts = await partsOf(
      await Document.toBuffer({
        body: [Build.table([Build.row([Build.cell("x", { shading })])], { shading })],
        styles: [style]
      })
    );
    const shd = (xml: string) => xml.match(/<w:shd [^>]*\/>/g) ?? [];
    const written = '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>';
    expect(shd(parts.get("word/styles.xml")!)).toEqual([written, written]);
    expect(shd(parts.get("word/document.xml")!)).toEqual([written, written]);
  });

  it("refuses a pattern outside the vocabulary, as a table does", async () => {
    const bad: StyleDef = {
      ...style,
      tableStyleConditions: undefined,
      tableProperties: { shading: { fill: "FFFFFF", pattern: "plaid" as never } }
    };
    await expect(Document.toBuffer({ body: [], styles: [bad] })).rejects.toThrow(DocxArgumentError);
  });
});

describe("the DOCX order check reaches", () => {
  const NS = [
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    'xmlns:v="urn:schemas-microsoft-com:vml"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
  ].join(" ");
  /** A paragraph whose properties are out of order: `numPr` belongs before `spacing`. */
  const BAD_P = "<w:p><w:pPr><w:spacing/><w:numPr/></w:pPr></w:p>";
  const BAD_P_REPORT = "child <numPr> appears after <spacing>";
  const run = (inner: string) => `<w:p><w:r>${inner}</w:r></w:p>`;
  const anchored = (graphicData: string) =>
    run(
      `<w:drawing><wp:anchor><a:graphic><a:graphicData>${graphicData}</a:graphicData></a:graphic></wp:anchor></w:drawing>`
    );
  const vml = `<w:pict><v:shape><v:textbox><w:txbxContent>${BAD_P}</w:txbxContent></v:textbox></v:shape></w:pict>`;
  const change = (name: string, inner: string) =>
    `<w:${name} w:id="1" w:author="A">${inner}</w:${name}>`;

  const cases: [where: string, body: string, trail: string, report: string][] = [
    [
      "runs in a block-level revision",
      "<w:ins><w:r><w:t/><w:rPr/></w:r></w:ins>",
      "document/body/ins/r",
      "child <rPr> appears after <t>"
    ],
    [
      "a DrawingML shape",
      anchored("<wps:wsp><wps:bodyPr/><wps:spPr/></wps:wsp>"),
      "graphicData/wsp",
      "child <spPr> appears after <bodyPr>"
    ],
    [
      "a DrawingML text box's content",
      anchored(`<wps:wsp><wps:txbx><w:txbxContent>${BAD_P}</w:txbxContent></wps:txbx></wps:wsp>`),
      "txbx/txbxContent/p/pPr",
      BAD_P_REPORT
    ],
    ["a VML text box's content", run(vml), "textbox/txbxContent/p/pPr", BAD_P_REPORT],
    [
      "an alternative of markup compatibility",
      run(`<mc:AlternateContent><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent>`),
      "r/AlternateContent/Fallback/pict/shape/textbox/txbxContent/p/pPr",
      BAD_P_REPORT
    ],
    [
      "ruby",
      run("<w:ruby><w:rubyBase/><w:rt/></w:ruby>"),
      "r/ruby",
      "child <rt> appears after <rubyBase>"
    ],
    [
      "the runs of a ruby's base text",
      run("<w:ruby><w:rubyBase><w:r><w:t/><w:rPr/></w:r></w:rubyBase></w:ruby>"),
      "ruby/rubyBase/r",
      "child <rPr> appears after <t>"
    ],
    [
      "a cell's previous properties",
      `<w:tbl><w:tr><w:tc><w:tcPr>${change("tcPrChange", "<w:tcPr><w:shd/><w:tcW/></w:tcPr>")}</w:tcPr></w:tc></w:tr></w:tbl>`,
      "tcPrChange/tcPr",
      "child <tcW> appears after <shd>"
    ],
    [
      "a table's previous properties",
      `<w:tbl><w:tblPr>${change("tblPrChange", "<w:tblPr><w:tblLook/><w:tblStyle/></w:tblPr>")}</w:tblPr></w:tbl>`,
      "tblPrChange/tblPr",
      "child <tblStyle> appears after <tblLook>"
    ],
    [
      "a section's previous properties",
      `<w:sectPr>${change("sectPrChange", "<w:sectPr><w:pgMar/><w:pgSz/></w:sectPr>")}</w:sectPr>`,
      "sectPrChange/sectPr",
      "child <pgSz> appears after <pgMar>"
    ],
    [
      "a paragraph mark's previous run properties",
      `<w:p><w:pPr><w:rPr>${change("rPrChange", "<w:rPr><w:del/><w:ins/></w:rPr>")}</w:rPr></w:pPr></w:p>`,
      "rPrChange/rPr",
      "child <ins> appears after <del>"
    ],
    [
      "a text box in a drawing group",
      anchored(
        `<wpg:wgp><wpg:grpSp><wps:wsp><wps:txbx><w:txbxContent>${BAD_P}</w:txbxContent></wps:txbx></wps:wsp></wpg:grpSp></wpg:wgp>`
      ),
      "wgp/grpSp/wsp/txbx/txbxContent/p/pPr",
      BAD_P_REPORT
    ],
    [
      "a text box in a VML group, and in an embedded object",
      run(
        `<w:object><v:group><v:rect><v:textbox><w:txbxContent>${BAD_P}</w:txbxContent></v:textbox></v:rect></v:group></w:object>`
      ),
      "object/group/rect/textbox/txbxContent/p/pPr",
      BAD_P_REPORT
    ],
    [
      "the borders of a section's previous properties",
      `<w:sectPr>${change("sectPrChange", '<w:sectPr><w:pgBorders><w:bottom w:val="single"/><w:top w:val="single"/></w:pgBorders></w:sectPr>')}</w:sectPr>`,
      "sectPrChange/sectPr/pgBorders",
      "child <top> appears after <bottom>"
    ],
    [
      "the borders of a cell's previous properties",
      `<w:tbl><w:tr><w:tc><w:tcPr>${change("tcPrChange", '<w:tcPr><w:tcBorders><w:bottom w:val="single"/><w:top w:val="single"/></w:tcBorders></w:tcPr>')}</w:tcPr></w:tc></w:tr></w:tbl>`,
      "tcPrChange/tcPr/tcBorders",
      "child <top> appears after <bottom>"
    ]
  ];

  for (const [where, body, trail, report] of cases) {
    it(where, () => {
      const xml = `<w:document ${NS}><w:body>${body}</w:body></w:document>`;
      const messages = checkDocxElementOrder(new Map([["word/document.xml", xml]])).map(
        violation => violation.message
      );
      expect(messages).toEqual([expect.stringContaining(report)]);
      expect(messages[0]).toContain(`${trail}> child`);
    });
  }

  it("a style's previous paragraph properties", () => {
    const xml = `<w:styles ${NS}><w:style><w:pPr><w:pPrChange w:id="1" w:author="A"><w:pPr><w:spacing/><w:keepNext/></w:pPr></w:pPrChange></w:pPr></w:style></w:styles>`;
    expect(checkDocxElementOrder(new Map([["word/styles.xml", xml]]))).toEqual([
      {
        part: "word/styles.xml",
        message: expect.stringContaining("<styles/style/pPr/pPrChange/pPr> child <keepNext>")
      }
    ]);
  });

  it("a style's paragraph properties, as CT_PPrGeneral", () => {
    const xml = `<w:styles ${NS}><w:style><w:pPr><w:spacing/><w:keepNext/></w:pPr></w:style></w:styles>`;
    expect(checkDocxElementOrder(new Map([["word/styles.xml", xml]]))).toEqual([
      {
        part: "word/styles.xml",
        message: expect.stringContaining("PPrGeneral requires")
      }
    ]);
  });
});

describe("change records where the schema has none", () => {
  const revision = { id: 3, author: "A" };
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  it("sets aside, as not written back, a change record inside a change record's previous properties", () => {
    const tcPr = parseXml(
      `<w:tcPr ${W}><w:tcPrChange w:id="1" w:author="A"><w:tcPr><w:vAlign w:val="top"/><w:tcPrChange w:id="2" w:author="B"><w:tcPr/></w:tcPrChange></w:tcPr></w:tcPrChange></w:tcPr>`
    ).root;
    const cell = parseTableCellProperties(tcPr);
    // The previous properties are read — they used to be dropped — and the nested change is set aside.
    expect(cell.propertyChange?.previousProperties?.verticalAlign).toBe("top");
    expect(cell.propertyChange?.previousProperties?.propertyChange).toBeUndefined();
    const doc: DocxDocument = {
      body: [
        Build.table([Build.row([{ ...Build.cell("x"), properties: cell }])], {
          columnWidths: [100]
        })
      ]
    };
    expect(Document.listUnrecognizedValues(doc)).toEqual([
      expect.objectContaining({ writtenBack: false, value: expect.stringContaining("tcPrChange") })
    ]);

    const trPr = parseXml(
      `<w:trPr ${W}><w:trPrChange w:id="1" w:author="A"><w:trPr><w:cantSplit/><w:trPrChange w:id="2" w:author="B"><w:trPr/></w:trPrChange></w:trPr></w:trPrChange></w:trPr>`
    ).root;
    const row = parseTableRowProperties(trPr);
    expect(row.propertyChange?.previousProperties?.cantSplit).toBe(true);
    expect(row.propertyChange?.previousProperties?.propertyChange).toBeUndefined();

    const tblPr = parseXml(
      `<w:tblPr ${W}><w:tblPrChange w:id="1" w:author="A"><w:tblPr><w:tblStyle w:val="X"/><w:tblPrChange w:id="2" w:author="B"><w:tblPr/></w:tblPrChange></w:tblPr></w:tblPrChange></w:tblPr>`
    ).root;
    const table = parseTableProperties(tblPr);
    expect(table.propertyChange?.previousProperties?.style).toBe("X");
    expect(table.propertyChange?.previousProperties?.propertyChange).toBeUndefined();
  });

  it("refuses one a caller sets there, rather than dropping it — when built and when written", async () => {
    const nested = { revision, previousProperties: {} };
    const previous = { propertyChange: nested };
    const change = { revision, previousProperties: previous };
    expect(() => Build.table([Build.row([Build.cell("x")])], { propertyChange: change })).toThrow(
      DocxArgumentError
    );
    expect(() => Build.row([Build.cell("x")], { propertyChange: change })).toThrow(
      DocxArgumentError
    );
    expect(() => Build.cell("x", { propertyChange: change })).toThrow(DocxArgumentError);

    // A model built without the builders reaches the writer, which refuses it too.
    const cell = { content: [Build.textParagraph("x")] };
    const tables: Table[] = [
      { type: "table", properties: { propertyChange: change }, rows: [{ cells: [cell] }] },
      { type: "table", rows: [{ properties: { propertyChange: change }, cells: [cell] }] },
      { type: "table", rows: [{ cells: [{ ...cell, properties: { propertyChange: change } }] }] }
    ];
    for (const table of tables) {
      await expect(Document.toBuffer({ body: [table] })).rejects.toThrow(DocxArgumentError);
    }
  });

  it("writes a cell's previous properties without a placeholder width", async () => {
    const document = await partsOf(
      await Document.toBuffer({
        body: [
          Build.table(
            [
              Build.row([
                Build.cell("x", {
                  propertyChange: { revision, previousProperties: { verticalAlign: "top" } }
                })
              ])
            ],
            { columnWidths: [100] }
          )
        ]
      })
    );
    const change = document
      .get("word/document.xml")!
      .match(/<w:tcPrChange[\s\S]*?<\/w:tcPrChange>/)![0];
    expect(change).toContain('<w:vAlign w:val="top"/>');
    expect(change).not.toContain("w:tcW");
  });
});
