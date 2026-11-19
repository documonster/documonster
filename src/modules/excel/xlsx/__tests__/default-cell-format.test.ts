import { zip } from "@archive/create-archive";
import { extractAll } from "@archive/unzip/extract";
import { cellGetValue, type CellData } from "@excel/core/cell";
import { Cell, Workbook } from "@excel/index";
import { WorkbookReader } from "@excel/stream/workbook-reader";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { describe, expect, it } from "vitest";

/**
 * `cellXfs[0]` is the format of every cell with no `s` (ECMA-376 §18.3.1.4), and of every cell the sheet does not
 * hold. Excel writes it as General in the default font — which sets nothing — until the Normal style is changed. A
 * cell without `s` used to read as unformatted whatever it said, and the writer replaced `cellXfs[0]` with its own.
 */

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** fonts: 0 Calibri, 1 Arial bold. cellXfs: 0 is `0.00` in Arial bold (as is the Normal style), 1 plain, 2 `m/d/yy`. */
const CHANGED_DEFAULT = `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="2" fontId="1" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="1" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs>`;

async function workbookWith(styles: string, sheetData: string): Promise<Uint8Array> {
  const archive = zip();
  const files: Record<string, string> = {
    "[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": `<?xml version="1.0"?><styleSheet xmlns="${MAIN}">${styles}</styleSheet>`,
    "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="${MAIN}"><sheetData>${sheetData}</sheetData></worksheet>`
  };
  for (const [name, text] of Object.entries(files)) {
    archive.add(name, new TextEncoder().encode(text));
  }
  return archive.toBuffer();
}

const CELLS =
  '<row r="1"><c r="A1"><v>1.5</v></c><c r="B1" s="0"><v>2.5</v></c><c r="C1" s="1"><v>3.5</v></c>' +
  '<c r="D1" s="3"><v>0.25</v></c></row>';

const ARIAL = { bold: true, size: 14, name: "Arial" };

/** What the written styles part says: each written cell's xf as `numFmtId/fontId`, and the two `[0]` xfs. */
async function written(bytes: Uint8Array): Promise<{
  readonly cells: Record<string, string>;
  readonly defaultXf: string;
  readonly normalXf: string;
}> {
  const parts = await extractAll(bytes);
  const text = (name: string) => new TextDecoder().decode(parts.get(name)!.data);
  const styles = text("xl/styles.xml");
  const fonts = [...styles.matchAll(/<font>(.*?)<\/font>/gs)].map(m =>
    /Arial/.test(m[1]!) ? "Arial" : "Calibri"
  );
  const xfs = (tag: string) =>
    [
      ...new RegExp(`<${tag}[^>]*>(.*?)</${tag}>`, "s")
        .exec(styles)![1]!
        .matchAll(/<xf numFmtId="(\d+)" fontId="(\d+)"/g)
    ].map(m => `${m[1]}/${fonts[Number(m[2])]}`);
  const cellXfs = xfs("cellXfs");
  const cells: Record<string, string> = {};
  for (const [, r, s] of text("xl/worksheets/sheet1.xml").matchAll(
    /<c r="(\w+)"(?: s="(\d+)")?/g
  )) {
    cells[r!] = cellXfs[Number(s ?? 0)]!;
  }
  return { cells, defaultXf: cellXfs[0]!, normalXf: xfs("cellStyleXfs")[0]! };
}

describe("the default cell format", () => {
  it("is the workbook's, a cell with no s has no style of its own, and a plain format is stated", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    const ws = Workbook.getWorksheet(wb, "S")!;
    expect(Workbook.getDefaultStyle(wb)).toEqual({ numFmt: "0.00", font: ARIAL });
    expect(Cell.findStyle(ws, "A1")).toEqual({});
    expect(Cell.findStyle(ws, "B1")).toMatchObject({ numFmt: "0.00", font: ARIAL });
    expect(Cell.findStyle(ws, "C1")).toMatchObject({
      numFmt: "General",
      font: { size: 11, name: "Calibri" }
    });
  });

  it("is written back, with the Normal style's, so cells keep their format", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", 9);
    const out = await written(await Workbook.toBuffer(wb));
    expect(out.defaultXf).toBe("2/Arial");
    expect(out.normalXf).toBe("2/Arial");
    // A1 and E1 — a cell added since — have the default; B1 states it, and C1 its plain format.
    expect(out.cells).toEqual({
      A1: "2/Arial",
      B1: "2/Arial",
      C1: "0/Calibri",
      D1: expect.stringMatching(/\/Arial$/),
      E1: "2/Arial"
    });
  });

  // A cell added since and one read without `s` are the same thing, and stay so through a round trip.
  it("reads the same before and after a round trip", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", 9);
    const back = await Workbook.read(await Workbook.toBuffer(wb));
    const again = Workbook.getWorksheet(back, "S")!;
    // The xf-level `apply*` flags a write adds are not compared: they say nothing about how a cell looks.
    for (const address of ["A1", "B1", "C1", "D1", "E1"]) {
      expect(Cell.getNumFmt(again, address)).toEqual(Cell.getNumFmt(ws, address));
      expect(Cell.getFont(again, address)).toEqual(Cell.getFont(ws, address));
    }
    expect([Cell.findStyle(again, "A1"), Cell.findStyle(again, "E1")]).toEqual([{}, {}]);
    expect(Workbook.getDefaultStyle(back)).toEqual(Workbook.getDefaultStyle(wb));
  });

  // A style that sets some facets has the default's for the rest; only a facet set to its plain value overrides it.
  it.each(["xlsx", "xlsb"] as const)(
    "keeps the default's facets a style leaves to it, through %s",
    async format => {
      const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
      const ws = Workbook.getWorksheet(wb, "S")!;
      Cell.setFont(ws, "A1", { name: "Arial", size: 10 });
      Cell.setNumFmt(ws, "E1", "General");
      Cell.setValue(ws, "E1", 9);
      const back = Workbook.getWorksheet(
        await Workbook.read(await Workbook.toBuffer(wb, { format }), { format }),
        "S"
      )!;
      expect(Cell.getNumFmt(back, "A1")).toBe("0.00");
      expect(Cell.getFont(back, "A1")).toMatchObject({ name: "Arial", size: 10 });
      expect(Cell.getNumFmt(back, "E1")).toBe("General");
      expect(Cell.getFont(back, "E1")).toMatchObject(ARIAL);
    }
  );

  it("is handed out as a copy", async () => {
    const styles = CHANGED_DEFAULT.replace(
      '<cellXfs count="4"><xf numFmtId="2"',
      '<cellXfs count="4"><xf numFmtId="7"'
    );
    const wb = await Workbook.read(await workbookWith(styles, CELLS));
    const copy = Workbook.getDefaultStyle(wb) as { font: { bold?: boolean }; _unrecognized: any };
    copy.font.bold = false;
    copy._unrecognized.numFmt.value = "8";
    expect(Workbook.getDefaultStyle(wb)).toMatchObject({ font: ARIAL });
    expect(Workbook.listUnrecognizedValues(wb)[0]).toEqual({
      path: "defaultStyle.numFmt",
      value: "7"
    });
  });

  // `Cell.getNumFmt` answers with the cell's own; what the cell shows is the default's where it has none.
  it("is what a cell with no number format of its own is shown in", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    const ws = Workbook.getWorksheet(wb, "S")!;
    Cell.setValue(ws, "E1", 9);
    expect(Cell.getNumFmt(ws, "A1")).toBeUndefined();
    expect(Cell.getDisplayText(ws, "A1")).toBe("1.50");
    expect(Cell.getDisplayText(ws, "E1")).toBe("9.00");
    expect(Cell.getDisplayText(ws, "C1")).toBe("3.5");
  });

  it("makes a number with no s a date when it is a date format", async () => {
    const styles = CHANGED_DEFAULT.replace(
      '<cellXfs count="4"><xf numFmtId="2"',
      '<cellXfs count="4"><xf numFmtId="14"'
    );
    const bytes = await workbookWith(styles, '<row r="1"><c r="A1"><v>45000</v></c></row>');
    const date = new Date(Date.UTC(2023, 2, 15));
    expect(Cell.getValue(Workbook.getWorksheet(await Workbook.read(bytes), "S")!, "A1")).toEqual(
      date
    );
    const values: unknown[] = [];
    for await (const ws of new WorkbookReader(bytes, { worksheets: "emit", styles: "cache" })) {
      for await (const row of ws) {
        for (const cell of (row as { cells: (CellData | undefined)[] }).cells) {
          if (cell) {
            values.push(cellGetValue(cell));
          }
        }
      }
    }
    expect(values).toEqual([date]);
  });

  it("survives a round trip through XLSB", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    const back = await Workbook.read(await Workbook.toBuffer(wb, { format: "xlsb" }), {
      format: "xlsb"
    });
    const ws = Workbook.getWorksheet(back, "S")!;
    expect(Workbook.getDefaultStyle(back)).toEqual({ numFmt: "0.00", font: ARIAL });
    expect(Cell.findStyle(ws, "A1")).toEqual({});
    expect(Cell.findStyle(ws, "B1")).toMatchObject({ numFmt: "0.00", font: ARIAL });
    expect(Cell.findStyle(ws, "C1")).toMatchObject({ numFmt: "General" });
    const out = await written(await Workbook.toBuffer(back));
    expect(out.defaultXf).toBe("2/Arial");
    expect(out.normalXf).toBe("2/Arial");
    expect(out.cells).toMatchObject({ A1: "2/Arial", B1: "2/Arial", C1: "0/Calibri" });
  });

  it("with a number format kept by its id, is reported, and dropped by discardUnrecognizedValues", async () => {
    const styles = CHANGED_DEFAULT.replace(
      '<cellXfs count="4"><xf numFmtId="2"',
      '<cellXfs count="4"><xf numFmtId="7"'
    );
    const wb = await Workbook.read(await workbookWith(styles, CELLS));
    // B1 names `cellXfs[0]` with `s="0"`, and so has the format as its own.
    expect(Workbook.listUnrecognizedValues(wb)).toEqual([
      { path: "defaultStyle.numFmt", value: "7" },
      { sheet: "S", path: "B1.style.numFmt", value: "7" }
    ]);
    expect(Workbook.discardUnrecognizedValues(wb)).toBe(2);
    expect(Workbook.getDefaultStyle(wb)).toEqual({ font: ARIAL });
    Cell.setValue(Workbook.getWorksheet(wb, "S")!, "E1", 9);
    const bytes = await Workbook.toBuffer(wb);
    expect((await written(bytes)).defaultXf).toBe("0/Arial");
    expect(Workbook.listUnrecognizedValues(await Workbook.read(bytes))).toEqual([]);
  });

  it.each(["xlsx", "xlsb"] as const)(
    "states General where a kept id is dropped from a format the default would give one, through %s",
    async format => {
      const styles = CHANGED_DEFAULT.replace(
        '<cellXfs count="4"><xf numFmtId="2"',
        '<cellXfs count="4"><xf numFmtId="7"'
      );
      const wb = await Workbook.read(await workbookWith(styles, CELLS));
      Workbook.discardUnrecognizedValues(wb, { filter: v => v.path === "B1.style.numFmt" });
      expect(Cell.getNumFmt(Workbook.getWorksheet(wb, "S")!, "B1")).toBe("General");
      const back = await Workbook.read(await Workbook.toBuffer(wb, { format }), { format });
      expect(Cell.getNumFmt(Workbook.getWorksheet(back, "S")!, "B1")).toBe("General");
      expect(Workbook.listUnrecognizedValues(back)).toEqual([
        { path: "defaultStyle.numFmt", value: "7" }
      ]);
    }
  );

  // A cell format takes, from its cell style, each facet whose `apply*` it does not set; an index of 0 is General, the
  // default font — a value, and one the Normal style here does not have.
  it("says a format applies the plain facets it states against a Normal style that sets them", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    Cell.setNumFmt(Workbook.getWorksheet(wb, "S")!, "E1", "General");
    Cell.setValue(Workbook.getWorksheet(wb, "S")!, "E1", 9);
    const parts = await extractAll(await Workbook.toBuffer(wb));
    const sheetXml = new TextDecoder().decode(parts.get("xl/worksheets/sheet1.xml")!.data);
    const stylesXml = new TextDecoder().decode(parts.get("xl/styles.xml")!.data);
    const s = Number(/<c r="E1" s="(\d+)"/.exec(sheetXml)![1]);
    const xf = [.../<cellXfs[^>]*>(.*?)<\/cellXfs>/s.exec(stylesXml)![1]!.matchAll(/<xf [^>]*>/g)][
      s
    ]![0];
    expect(xf).toContain('numFmtId="0"');
    expect(xf).toContain('applyNumberFormat="1"');
  });

  // [MS-OI29500] §18.8.45: in `cellXfs` an `apply*` says whether a later change to the cell style reaches the cell, not
  // which facet it shows — Excel shows a cell format's own number format with the flag absent, `"0"` or `"1"`. So a read
  // takes the format's own, and a write says so wherever it differs from its cell style's, for the readers that do not.
  it.each([
    ["absent", ""],
    ['"0"', ' applyNumberFormat="0"'],
    ['"1"', ' applyNumberFormat="1"']
  ])(
    "reads a format's own number format, whatever its cell style's, with applyNumberFormat %s",
    async (_flag, attribute) => {
      const styles = CHANGED_DEFAULT.replace(
        '<cellStyleXfs count="1"><xf numFmtId="2"',
        '<cellStyleXfs count="1"><xf numFmtId="14"'
      ).replace(
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
        `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"${attribute}/>`
      );
      const bytes = await workbookWith(styles, '<row r="1"><c r="A1" s="1"><v>45000</v></c></row>');
      expect(Cell.getValue(Workbook.getWorksheet(await Workbook.read(bytes), "S")!, "A1")).toBe(
        45000
      );
      for await (const ws of new WorkbookReader(bytes, { worksheets: "emit", styles: "cache" })) {
        for await (const row of ws) {
          const cell = (row as { cells: (CellData | undefined)[] }).cells[0]!;
          expect(cellGetValue(cell)).toBe(45000);
        }
      }
    }
  );

  it("says a format applies each facet its cell style holds otherwise, in both containers", async () => {
    const wb = await Workbook.read(await workbookWith(CHANGED_DEFAULT, CELLS));
    // C1 is General in the first font; the Normal style is `0.00` in Arial.
    const parts = await extractAll(await Workbook.toBuffer(wb));
    const sheetXml = new TextDecoder().decode(parts.get("xl/worksheets/sheet1.xml")!.data);
    const stylesXml = new TextDecoder().decode(parts.get("xl/styles.xml")!.data);
    const s = Number(/<c r="C1" s="(\d+)"/.exec(sheetXml)![1]);
    const xf = [.../<cellXfs[^>]*>(.*?)<\/cellXfs>/s.exec(stylesXml)![1]!.matchAll(/<xf [^>]*>/g)][
      s
    ]![0];
    expect(xf).toMatch(/numFmtId="0" fontId="0"/);
    expect(xf).toContain('applyNumberFormat="1"');
    expect(xf).toContain('applyFont="1"');

    const binary = await extractAll(await Workbook.toBuffer(wb, { format: "xlsb" }));
    const formats: { formatId: number; overrides: number }[] = [];
    let inCellXfs = false;
    for (const record of iterateInterpretableRecords(
      binary.get("xl/styles.bin")!.data,
      "xl/styles.bin"
    )) {
      const name = recordSpec(record.id)?.name;
      inCellXfs = name === "BrtBeginCellXFs" ? true : name === "BrtEndCellXFs" ? false : inCellXfs;
      if (inCellXfs && name === "BrtXF") {
        formats.push({
          formatId: record.payload[2]! + 256 * record.payload[3]!,
          overrides: record.payload[14]!
        });
      }
    }
    // Every cell format at `iFmt` 0 differs from the Normal style's `0.00`, and so overrides it.
    const general = formats.filter(f => f.formatId === 0);
    expect(general.length).toBeGreaterThan(0);
    for (const format of general) {
      expect(format.overrides & 0x01).toBe(0x01);
    }
  });

  it("does not let an edit to one cell's live style reach another's kept values", async () => {
    const styles = CHANGED_DEFAULT.replace(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '<xf numFmtId="7" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    );
    const wb = await Workbook.read(
      await workbookWith(
        styles,
        '<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>'
      )
    );
    const ws = Workbook.getWorksheet(wb, "S")!;
    (
      Cell.getStyle(ws, "A1") as { _unrecognized: { numFmt: { value: string } } }
    )._unrecognized.numFmt.value = "8";
    expect(Workbook.listUnrecognizedValues(wb)).toContainEqual({
      sheet: "S",
      path: "B1.style.numFmt",
      value: "7"
    });
  });

  it("gives a date with no format of its own the default's date format, in both containers", async () => {
    const styles = CHANGED_DEFAULT.replace(
      '<cellXfs count="4"><xf numFmtId="2"',
      '<cellXfs count="4"><xf numFmtId="22"'
    );
    const wb = await Workbook.read(await workbookWith(styles, CELLS));
    Cell.setValue(Workbook.getWorksheet(wb, "S")!, "E1", new Date(Date.UTC(2023, 2, 15, 12, 15)));
    for (const format of ["xlsx", "xlsb"] as const) {
      const back = await Workbook.read(await Workbook.toBuffer(wb, { format }), { format });
      expect(Cell.getNumFmt(Workbook.getWorksheet(back, "S")!, "E1")).toBe("m/d/yy h:mm");
    }
  });

  it.each([
    ["Excel's", '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'],
    [
      "LibreOffice's, which restates the defaults",
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyFont="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf>'
    ]
  ])("sets nothing when it is %s", async (_who, xf) => {
    const styles = `<numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1">${xf}</cellXfs>`;
    const wb = await Workbook.read(
      await workbookWith(styles, '<row r="1"><c r="A1"><v>1</v></c></row>')
    );
    expect(Workbook.getDefaultStyle(wb)).toBeUndefined();
    expect(Cell.findStyle(Workbook.getWorksheet(wb, "S")!, "A1")).toEqual({});
    expect((await written(await Workbook.toBuffer(wb))).defaultXf).toBe("0/Calibri");
  });
});
