import { readFileSync } from "fs";

import { extractAll } from "@archive/unzip/extract";
import { createZipSync } from "@archive/zip/zip-bytes";
import { Cell, Workbook, Worksheet } from "@excel/index";
import { makeTestDataPath } from "@test/utils";
import { describe, expect, it } from "vitest";

/**
 * The `<dxfs>` table survives a round trip at its source indices.
 *
 * A `dxfId` is an index, and not every holder of one is modelled: a pivot table's `<formats>` and a
 * `<colorFilter>` are preserved as XML with the number they were read with. The writer used to rebuild the
 * table from only the styles it could name, so those numbers pointed past its end (#237 — Excel repairs the
 * pivot) or, once anything else was allocated, silently at a *different* format.
 */
const dataPath = makeTestDataPath(import.meta.url, "../../../../__tests__/data");
const decoder = new TextDecoder();
const encoder = new TextEncoder();

type WorkbookHandle = ReturnType<typeof Workbook.create>;

/** #237's attachment: dxf 0 is a red font, dxf 1 bold; the pivot's formats reference 1 then 0. */
const PIVOT_FIXTURE = new Uint8Array(readFileSync(dataPath("pivot-format-dxf.xlsx")));

async function text(bytes: Uint8Array, part: string): Promise<string> {
  const entry = (await extractAll(bytes)).get(part);
  if (!entry) {
    throw new Error(`missing ${part}`);
  }
  return decoder.decode(entry.data);
}

async function patchPart(
  bytes: Uint8Array,
  part: string,
  patch: (xml: string) => string
): Promise<Uint8Array> {
  const entries = await extractAll(bytes);
  return createZipSync(
    [...entries.values()].map(entry => ({
      name: entry.path,
      data: entry.path === part ? encoder.encode(patch(decoder.decode(entry.data))) : entry.data
    }))
  );
}

/** Each `<dxf>` body, in index order — an empty `<dxf/>` included, since it still takes an index. */
async function dxfTable(bytes: Uint8Array): Promise<string[]> {
  const styles = await text(bytes, "xl/styles.xml");
  const block = /<dxfs\b[^>]*?(?:\/>|>[\s\S]*?<\/dxfs>)/.exec(styles)?.[0] ?? "";
  const bodies = [...block.matchAll(/<dxf\s*\/>|<dxf>([\s\S]*?)<\/dxf>/g)].map(m => m[1] ?? "");
  // The declared count has to agree with what follows, or Excel repairs the part.
  expect(Number(/<dxfs\b[^>]*?count="(\d+)"/.exec(block)?.[1] ?? 0)).toBe(bodies.length);
  return bodies;
}

/** A readable name for a dxf body, so assertions say what a reference *means*. */
function describeDxf(body: string | undefined): string {
  if (body === undefined) {
    return "DANGLING";
  }
  if (/<b\/>/.test(body)) {
    return "bold";
  }
  if (/<i\/>/.test(body)) {
    return "italic";
  }
  if (/FFFF0000/.test(body)) {
    return "red";
  }
  if (/FF00FF00/.test(body)) {
    return "green-fill";
  }
  return body;
}

/** What every `dxfId="n"` in `part` resolves to, in document order. */
async function resolve(bytes: Uint8Array, part: string, element: string): Promise<string[]> {
  const table = await dxfTable(bytes);
  const xml = await text(bytes, part);
  const pattern = new RegExp(`<${element}\\b[^>]*?\\sdxfId="(\\d+)"`, "g");
  return [...xml.matchAll(pattern)].map(m => describeDxf(table[Number(m[1])]));
}

const pivotFormats = (bytes: Uint8Array): Promise<string[]> =>
  resolve(bytes, "xl/pivotTables/pivotTable1.xml", "format");

async function load(bytes: Uint8Array): Promise<WorkbookHandle> {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes);
  return wb;
}

describe("pivot table <format dxfId> (#237)", () => {
  it("fixture: the formats reference bold, then red", async () => {
    expect(await pivotFormats(PIVOT_FIXTURE)).toEqual(["bold", "red"]);
  });

  it("writes the source <dxfs> back unchanged, so every reference still means what it did", async () => {
    const out = await Workbook.toBuffer(await load(PIVOT_FIXTURE));
    expect(await dxfTable(out)).toEqual(await dxfTable(PIVOT_FIXTURE));
    expect(await pivotFormats(out)).toEqual(["bold", "red"]);
  });

  it("is stable across repeated writes and a re-read", async () => {
    const wb = await load(PIVOT_FIXTURE);
    const first = await Workbook.toBuffer(wb);
    const second = await Workbook.toBuffer(wb);
    expect(await dxfTable(second)).toEqual(await dxfTable(first));
    const reread = await Workbook.toBuffer(await load(first));
    expect(await dxfTable(reread)).toEqual(await dxfTable(PIVOT_FIXTURE));
    expect(await pivotFormats(reread)).toEqual(["bold", "red"]);
  });

  it("appends a new conditional format after the source table rather than over it", async () => {
    const wb = await load(PIVOT_FIXTURE);
    Worksheet.addConditionalFormatting(Workbook.getWorksheet(wb, 1)!, {
      ref: "A1:A3",
      rules: [
        { type: "expression", formulae: ["TRUE"], priority: 1, style: { font: { italic: true } } }
      ]
    });
    const out = await Workbook.toBuffer(wb);
    expect((await dxfTable(out)).map(describeDxf)).toEqual(["red", "bold", "italic"]);
    expect(await pivotFormats(out)).toEqual(["bold", "red"]);
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "cfRule")).toEqual(["italic"]);
  });
});

describe("<colorFilter dxfId> and conditional formatting sharing one table", () => {
  /**
   * A workbook whose `<dxfs>` holds [red font, green fill] — the second only reachable through a
   * `<colorFilter>`, which is preserved XML the model cannot name — plus a rule using the first.
   */
  async function source(): Promise<Uint8Array> {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", "h");
    Cell.setValue(ws, "A2", 1);
    Worksheet.addConditionalFormatting(ws, {
      ref: "A2:A5",
      rules: [
        {
          type: "expression",
          formulae: ["TRUE"],
          priority: 1,
          style: { font: { color: { argb: "FFFF0000" } } }
        }
      ]
    });
    const written = await Workbook.toBuffer(wb);
    const withDxf = await patchPart(written, "xl/styles.xml", xml =>
      xml
        .replace(/<dxfs count="1">/, '<dxfs count="2">')
        .replace(
          /<\/dxfs>/,
          '<dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>'
        )
    );
    return patchPart(withDxf, "xl/worksheets/sheet1.xml", xml =>
      xml.replace(
        "</sheetData>",
        '</sheetData><autoFilter ref="A1:A5"><filterColumn colId="0"><colorFilter dxfId="1"/></filterColumn></autoFilter>'
      )
    );
  }

  it("fixture: the rule is red and the filter green", async () => {
    const bytes = await source();
    expect((await dxfTable(bytes)).map(describeDxf)).toEqual(["red", "green-fill"]);
    expect(await resolve(bytes, "xl/worksheets/sheet1.xml", "colorFilter")).toEqual(["green-fill"]);
  });

  it("keeps the filter on its own format and does not duplicate the rule's", async () => {
    const out = await Workbook.toBuffer(await load(await source()));
    // The rule read from the file resolves back to its own entry — no third, duplicate dxf.
    expect((await dxfTable(out)).map(describeDxf)).toEqual(["red", "green-fill"]);
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "cfRule")).toEqual(["red"]);
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "colorFilter")).toEqual(["green-fill"]);
  });

  it("survives XLSX → XLSB → XLSX with the filter still on its own format", async () => {
    const xlsb = await Workbook.toBuffer(await load(await source()), { format: "xlsb" });
    const out = await Workbook.toBuffer(await load(new Uint8Array(xlsb)));
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "cfRule")).toEqual(["red"]);
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "colorFilter")).toEqual(["green-fill"]);
  });
});

describe("useStyles: false", () => {
  // Turning off cell styles is not a licence to break references into `<dxfs>`, which is structural.
  it("keeps a read workbook's pivot formats pointing at their own formats", async () => {
    const out = await Workbook.toBuffer(await load(PIVOT_FIXTURE), { useStyles: false } as never);
    expect(await dxfTable(out)).toEqual(await dxfTable(PIVOT_FIXTURE));
    expect(await pivotFormats(out)).toEqual(["bold", "red"]);
  });

  it("writes a conditional format's style instead of throwing", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Worksheet.addConditionalFormatting(ws, {
      ref: "A1",
      rules: [
        {
          type: "expression",
          formulae: ["TRUE"],
          priority: 1,
          style: { font: { bold: true }, numFmt: "0.000%" }
        }
      ]
    });
    const out = await Workbook.toBuffer(wb, { useStyles: false } as never);
    expect(await resolve(out, "xl/worksheets/sheet1.xml", "cfRule")).toEqual(["bold"]);
    const styles = await text(out, "xl/styles.xml");
    const id = /<dxf>[\s\S]*?<numFmt numFmtId="(\d+)" formatCode="0.000%"/.exec(styles)?.[1];
    expect(id).toBeDefined();
    expect(styles).toMatch(
      new RegExp(`<numFmts[^>]*>[\\s\\S]*numFmtId="${id}" formatCode="0.000%"`)
    );
  });
});

describe("a differential number format of General", () => {
  it("is written, although its built-in id is 0", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Worksheet.addConditionalFormatting(ws, {
      ref: "A1",
      rules: [{ type: "expression", formulae: ["TRUE"], priority: 1, style: { numFmt: "General" } }]
    });
    const [body] = await dxfTable(await Workbook.toBuffer(wb));
    expect(body).toContain('<numFmt numFmtId="0" formatCode="General"');
  });
});

describe("a differential number format's id", () => {
  // `numFmtId` is an allocation of one write — it depends on what else that workbook registered first — so
  // it must not be written onto the caller's style, which is routinely shared between workbooks.
  function withRules(styles: Record<string, unknown>[]): WorkbookHandle {
    const wb = Workbook.create();
    Worksheet.addConditionalFormatting(Workbook.addWorksheet(wb, "S"), {
      ref: "A1",
      rules: styles.map((style, index) => ({
        type: "expression",
        formulae: ["TRUE"],
        priority: index + 1,
        style
      }))
    });
    return wb;
  }

  /** Every `numFmtId` a `<dxf>` names, and whether `<numFmts>` declares it with the same code. */
  async function dxfNumFmts(bytes: Uint8Array): Promise<string[]> {
    const styles = await text(bytes, "xl/styles.xml");
    const declared = /<numFmts\b[\s\S]*?<\/numFmts>/.exec(styles)?.[0] ?? "";
    return (await dxfTable(bytes)).flatMap(body =>
      [...body.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]+)"/g)].map(([, id, code]) =>
        declared.includes(`numFmtId="${id}" formatCode="${code}"`) ? `${code}` : `UNDECLARED ${id}`
      )
    );
  }

  it("is not written onto the caller's style", async () => {
    const shared = { numFmt: "0.000%" };
    await Workbook.toBuffer(withRules([shared]));
    expect(shared).toEqual({ numFmt: "0.000%" });
  });

  it("stays declared when two workbooks sharing one style are written concurrently", async () => {
    const shared = { numFmt: "0.000%" };
    const [plain, preceded] = await Promise.all([
      Workbook.toBuffer(withRules([shared])),
      Workbook.toBuffer(withRules([{ numFmt: "0.0000" }, shared]))
    ]);
    expect(await dxfNumFmts(plain)).toEqual(["0.000%"]);
    expect(await dxfNumFmts(preceded)).toEqual(["0.0000", "0.000%"]);
  });
});

describe("a differential format's content survives, not just its index", () => {
  /** Replace the `<dxfs>` of a small workbook with `dxfs`, and read it. */
  async function withDxfs(dxfs: string): Promise<Uint8Array> {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    return patchPart(await Workbook.toBuffer(wb), "xl/styles.xml", xml =>
      xml.replace(/<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/, dxfs)
    );
  }

  const roundTrip = async (source: Uint8Array): Promise<string[]> =>
    dxfTable(await Workbook.toBuffer(await load(source)));

  it("keeps Excel's preset highlight fill, which has no patternType", async () => {
    // Every "Light Red Fill" Excel writes. Read as `none` and written back with `patternType="none"`, it said
    // "remove the fill" — LibreOffice draws the first and not the second.
    const body = '<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>';
    expect(await roundTrip(await withDxfs(`<dxfs count="1"><dxf>${body}</dxf></dxfs>`))).toEqual([
      body
    ]);
  });

  it('reads <b val="0"/> as not bold, and writes it back', async () => {
    // Read as bold, which inverted a rule that removes bold into one that adds it.
    const body = '<font><b val="0"/><i/><strike val="false"/></font>';
    const out = await roundTrip(await withDxfs(`<dxfs count="1"><dxf>${body}</dxf></dxfs>`));
    expect(out).toEqual(['<font><b val="0"/><i/><strike val="0"/></font>']);
  });

  it("keeps a border's inner edges, and an unknown child no longer drops the whole border", async () => {
    const out = await roundTrip(
      await withDxfs(
        '<dxfs count="1"><dxf><border><left style="thin"/><vertical style="thin"/>' +
          '<horizontal style="hair"/><x:future xmlns:x="urn:x"><x:deep/></x:future></border></dxf></dxfs>'
      )
    );
    expect(out[0]).toContain('<left style="thin"/>');
    expect(out[0]).toContain('<vertical style="thin"/><horizontal style="hair"/>');
  });

  it("keeps the resets a differential format states", async () => {
    // Defaults a cell drops, and a differential format must not: each one resets the cell's own value.
    const body =
      '<font><vertAlign val="baseline"/></font>' +
      '<alignment horizontal="general" wrapText="0" indent="0" relativeIndent="-1" ' +
      'justifyLastLine="0" textRotation="0" readingOrder="0"/>';
    const [out] = await roundTrip(await withDxfs(`<dxfs count="1"><dxf>${body}</dxf></dxfs>`));
    expect(out).toContain('<vertAlign val="baseline"/>');
    for (const attribute of [
      'horizontal="general"',
      'wrapText="0"',
      'indent="0"',
      'relativeIndent="-1"',
      'justifyLastLine="0"',
      'textRotation="0"',
      'readingOrder="0"'
    ]) {
      expect(out).toContain(attribute);
    }
  });

  it("carries every facet through XLSX → XLSB → XLSX", async () => {
    const body =
      '<font><b val="0"/><u val="double"/><vertAlign val="subscript"/><sz val="12"/><color indexed="10"/></font>' +
      '<numFmt numFmtId="170" formatCode="0.0%"/>' +
      '<fill><patternFill patternType="lightGray"><fgColor theme="4"/></patternFill></fill>' +
      '<alignment horizontal="general" vertical="bottom" textRotation="135" indent="2" ' +
      'relativeIndent="3" readingOrder="0"/>' +
      '<border><left style="thin"/><horizontal style="hair"/></border>' +
      '<protection locked="0" hidden="1"/>';
    const source = await withDxfs(`<dxfs count="1"><dxf>${body}</dxf></dxfs>`);
    const direct = await roundTrip(source);
    const xlsb = await Workbook.toBuffer(await load(source), {
      format: "xlsb",
      unsupported: "error"
    } as never);
    expect(await roundTrip(new Uint8Array(xlsb))).toEqual(direct);
  });
});

describe("a differential format's <extLst>", () => {
  const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

  /** A workbook whose root declares `x14`, and whose `<dxfs>` is `dxfs`. */
  async function source(dxfs: string): Promise<Uint8Array> {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    return patchPart(await Workbook.toBuffer(wb), "xl/styles.xml", xml =>
      xml
        .replace("<styleSheet ", `<styleSheet xmlns:x14="${X14}" `)
        .replace(/<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/, dxfs)
    );
  }

  const EXT = '<extLst><ext uri="{A}"><x14:thing x14:val="1"/></ext></extLst>';

  it("is written back, last, and with the prefix it borrowed from the source's root declared", async () => {
    const out = await Workbook.toBuffer(
      await load(
        await source(`<dxfs count="2"><dxf><font><b/></font>${EXT}</dxf><dxf>${EXT}</dxf></dxfs>`)
      )
    );
    const [withFont, alone] = await dxfTable(out);
    const declared = `<extLst xmlns:x14="${X14}"><ext uri="{A}"><x14:thing x14:val="1"/></ext></extLst>`;
    expect(withFont).toBe(`<font><b/></font>${declared}`);
    expect(alone).toBe(declared);
  });

  it("survives a rule's format being changed, and a second write", async () => {
    const wb = await load(
      await source(`<dxfs count="1"><dxf><font><b/></font>${EXT}</dxf></dxfs>`)
    );
    const style = (wb as unknown as { _dxfs: { font?: { italic?: boolean } }[] })._dxfs[0]!;
    style.font = { ...style.font, italic: true };
    await Workbook.toBuffer(wb);
    const [body] = await dxfTable(await Workbook.toBuffer(wb));
    expect(body).toMatch(/^<font><b\/><i\/><\/font><extLst xmlns:x14=[^>]+><ext uri="\{A\}">/);
  });

  it("is reported when written as XLSB, which has nowhere to put it", async () => {
    const wb = await load(
      await source(`<dxfs count="1"><dxf><font><b/></font>${EXT}</dxf></dxfs>`)
    );
    await expect(
      Workbook.toBuffer(wb, { format: "xlsb", unsupported: "error" } as never)
    ).rejects.toThrow(/differential format 0: extLst/);
  });

  it("does not let an unknown element's children be read as facets", async () => {
    const out = await Workbook.toBuffer(
      await load(
        await source(
          '<dxfs count="1"><dxf><bogus><font><b/></font></bogus><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'
        )
      )
    );
    expect(await dxfTable(out)).toEqual([
      '<fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>'
    ]);
  });
});

describe("<dxfs> entries written as <mc:AlternateContent>", () => {
  it("counts each as one entry and reads its Fallback, so every later dxfId still lines up", async () => {
    // Hancom Office writes 8 of its 12 table-style formats as a `Requires="hs"` Choice plus a plain Fallback.
    // The first `<mc:…>` used to end the list: one format read, and every dxfId above 0 dangling.
    const source = new Uint8Array(readFileSync(dataPath("han-cell-namespace-prefixes.xlsx")));
    const out = await Workbook.toBuffer(await load(source));
    const table = await dxfTable(out);
    expect(table).toHaveLength(12);
    // Entry 0 is an AlternateContent; its Fallback carries a solid fill and four white edges.
    expect(table[0]).toContain('<patternFill patternType="solid"><fgColor rgb="ffd7dff4"/>');
    expect(table[0]).toContain('<vertical style="thin"><color rgb="ffffffff"/></vertical>');
  });

  it("keeps the entry's place when there is no Fallback", async () => {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    const source = await patchPart(await Workbook.toBuffer(wb), "xl/styles.xml", xml =>
      xml.replace(
        /<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/,
        '<dxfs count="2"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
          '<mc:Choice Requires="zz"><dxf><font><i/></font></dxf></mc:Choice></mc:AlternateContent>' +
          "<dxf><font><b/></font></dxf></dxfs>"
      )
    );
    expect(await dxfTable(await Workbook.toBuffer(await load(source)))).toEqual([
      "",
      "<font><b/></font>"
    ]);
  });
});
