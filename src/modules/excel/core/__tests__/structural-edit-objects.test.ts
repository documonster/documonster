/**
 * Structural edits, second part: what is positioned by sheet index or by drawing anchor, and what names a table the
 * edit removes. See `structural-edit.test.ts` for references in formulas and their holders.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import {
  Cell,
  Chart,
  Column,
  ConditionalFormatting,
  DefinedNames,
  Form,
  Image,
  Row,
  Table,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function scoped(wb: Workbook.Handle): Array<[string, number | undefined, string[]]> {
  return DefinedNames.list(Workbook.getDefinedNames(wb))
    .map(
      entry =>
        [entry.name, entry.localSheetId, entry.ranges] as [string, number | undefined, string[]]
    )
    .sort((a, b) => a[0].localeCompare(b[0]));
}

describe("sheet-scoped names on sheet removal", () => {
  function build() {
    const wb = Workbook.create();
    for (const name of ["A", "B", "C"]) {
      Workbook.addWorksheet(wb, name);
    }
    DefinedNames.setModel(Workbook.getDefinedNames(wb), [
      { name: "OnA", ranges: ["A!$A$1"], localSheetId: 0 },
      { name: "OnB", ranges: ["B!$A$1"], localSheetId: 1 },
      { name: "OnC", ranges: ["C!$A$1"], localSheetId: 2 },
      { name: "Twice", ranges: ["C!$B$1*2"], formulaExpression: "C!$B$1*2", localSheetId: 2 },
      { name: "Global", ranges: ["C!$A$1"] }
    ]);
    return wb;
  }

  it("removes the names scoped to the removed sheet and renumbers the rest", () => {
    const wb = build();
    Workbook.removeWorksheet(wb, "B");
    // A sheet named `C` is quoted, as the reference rewriter quotes it: unquoted it reads as an R1C1 column.
    expect(scoped(wb)).toEqual([
      ["Global", undefined, ["'C'!$A$1"]],
      ["OnA", 0, ["A!$A$1"]],
      ["OnC", 1, ["'C'!$A$1"]],
      ["Twice", 1, ["C!$B$1*2"]]
    ]);
  });

  it("keeps each name on its sheet through a write and a read", async () => {
    const wb = build();
    Workbook.removeWorksheet(wb, "A");
    for (const format of ["xlsx", "xlsb"] as const) {
      const back = await Workbook.read(await Workbook.toBuffer(wb, { format }));
      const names = scoped(back).filter(([name]) => name !== "Global");
      expect(names.map(([name, id]) => [name, id])).toEqual([
        ["OnB", 0],
        ["OnC", 1],
        ["Twice", 1]
      ]);
    }
  });

  it("renumbers names when a chartsheet before them is removed", () => {
    const wb = Workbook.create();
    const data = Workbook.addWorksheet(wb, "Data");
    Cell.setValue(data, "A1", 1);
    Workbook.addChartsheet(wb, "Chart", {
      chart: { type: "bar", series: [{ name: "x", values: "Data!$A$1:$A$1" }] }
    } as never);
    const last = Workbook.addWorksheet(wb, "Last");
    DefinedNames.setModel(Workbook.getDefinedNames(wb), [
      { name: "OnLast", ranges: ["Last!$A$1"], localSheetId: 2 }
    ]);
    expect(last).toBeDefined();
    Workbook.removeChartsheet(wb, "Chart");
    expect(scoped(wb)).toEqual([["OnLast", 1, ["Last!$A$1"]]]);
  });
});

describe("drawings and breaks follow inserted and deleted lines", () => {
  function sheet() {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    return { wb, ws };
  }

  it("moves a shape, keeping its size (a range-placed shape is move-but-don't-size)", () => {
    const { ws } = sheet();
    Image.addShape(ws, { type: "rect", range: "B3:D5" });
    Row.insert(ws, 1, []);
    Row.insert(ws, 5, []);
    Column.splice(ws, 1, 0, []);
    const range = Worksheet.getModel(ws).shapes![0].anchorRange!;
    expect([
      range.tl.nativeCol,
      range.tl.nativeRow,
      range.br!.nativeCol,
      range.br!.nativeRow
    ]).toEqual([2, 3, 5, 6]);
  });

  it("follows each picture's own placement", () => {
    const { wb, ws } = sheet();
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    const anchor = (col: number, row: number) => ({ col, row });
    Image.place(ws, id, { tl: anchor(0, 2), br: anchor(1, 6), editAs: "twoCell" });
    Image.place(ws, id, { tl: anchor(2, 2), br: anchor(3, 6), editAs: "oneCell" });
    Image.place(ws, id, { tl: anchor(4, 2), br: anchor(5, 6), editAs: "absolute" });
    Image.place(ws, id, { tl: anchor(6, 2), ext: { width: 10, height: 10 } });
    Row.insert(ws, 4, []); // strictly inside every picture
    const rows = Image.list(ws).map(image => [
      image.range!.tl.nativeRow,
      image.range!.br?.nativeRow
    ]);
    expect(rows).toEqual([
      [2, 7], //  move and size: grows
      [2, 6], //  move, don't size: the top did not move, nor does the size
      [2, 6], //  absolute: stays
      [2, undefined]
    ]);
    Row.insert(ws, 1, []); // above every picture
    expect(Image.list(ws).map(image => image.range!.tl.nativeRow)).toEqual([3, 3, 2, 3]);
  });

  it("puts a picture whose top row is deleted at the start of the band, not row 1", () => {
    const { wb, ws } = sheet();
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, id, { tl: { col: 0, row: 4 }, br: { col: 1, row: 8 }, editAs: "twoCell" });
    Row.splice(ws, 4, 3); // deletes rows 4–6 (0-based 3–5)
    const { tl, br } = Image.list(ws)[0].range!;
    expect([tl.nativeRow, tl.nativeRowOff, br!.nativeRow]).toEqual([3, 0, 5]);
  });

  it("moves a chart, which moves and sizes with cells by default", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", 1);
    Chart.addColumn(
      ws,
      { series: [{ name: "x", values: "S!$A$1:$A$1" }] },
      { tl: { col: 2, row: 2 }, br: { col: 6, row: 8 } }
    );
    Row.insert(ws, 5, []);
    const { tl, br } = Chart.get(ws)[0].range;
    expect([tl.nativeRow, br!.nativeRow]).toEqual([2, 9]);
  });

  it("moves a checkbox with the cell it is linked to, unless it is placed not to move", () => {
    const { ws } = sheet();
    const box = Form.addCheckbox(ws, "B3", { link: "C3" });
    const fixed = Form.addCheckbox(ws, "B5", { link: "C5", editAs: "absolute" });
    const sized = Form.addCheckbox(ws, "B7:C8", { editAs: "twoCell" });
    Row.insert(ws, 1, []);
    expect(box.model.link).toBe("$C$4");
    expect(box.model.tl.row).toBe(3);
    expect(box.model.br.row - box.model.tl.row).toBe(1);
    expect(fixed.model.link).toBe("$C$6");
    expect(fixed.model.tl.row).toBe(4);
    expect(sized.model.tl.row).toBe(7);
    // A row inserted inside a two-cell checkbox stretches it; one inside a one-cell checkbox does not.
    Row.insert(ws, 9, []);
    expect(sized.model.tl.row).toBe(7);
    expect(sized.model.br.row).toBe(9);
    const one = Form.addCheckbox(ws, "E2:F3");
    Row.insert(ws, 3, []);
    expect(one.model.tl.row).toBe(1);
    expect(one.model.br.row - one.model.tl.row).toBe(1);
  });

  it("moves page breaks and drops the ones in deleted lines", () => {
    const { ws } = sheet();
    Row.addPageBreak(ws, 5);
    Row.addPageBreak(ws, 10);
    Row.addPageBreak(ws, 20);
    Column.addPageBreak(ws, "C");
    Row.insert(ws, 2, []);
    Row.splice(ws, 12, 1); // the row below the second break (now 11 | 12)
    Column.splice(ws, 1, 0, []);
    expect(ws.rowBreaks.map(b => b.id)).toEqual([6, 20]);
    expect(ws.colBreaks.map(b => b.id)).toEqual([4]);
  });
});

describe("preserved legacy controls", () => {
  const CTRL_PROP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";

  /** Sheet S with a combo box at F4:H5 (0-based rows 3–4) linked to $E$1, input range $F$1:$F$3. */
  async function fixture(): Promise<Uint8Array> {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Form.addCheckbox(ws, "A1");
    const files = new Map<string, Uint8Array>();
    for (const [name, entry] of await extractAll(await Workbook.toBuffer(wb))) {
      files.set(name, entry.data);
    }
    const edit = (path: string, change: (text: string) => string) =>
      files.set(path, new TextEncoder().encode(change(new TextDecoder().decode(files.get(path)!))));
    files.set(
      "xl/ctrlProps/ctrlProp9.xml",
      new TextEncoder().encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Drop" dropStyle="combo" fmlaLink="$E$1" fmlaRange="$F$1:$F$3"/>`
      )
    );
    edit("xl/drawings/vmlDrawing1.vml", text =>
      text.replace(
        "</xml>",
        `<v:shape id="Drop_x0020_Down_x0020_2" o:spid="_x0000_s1030" type="#_x0000_t201" style="position:absolute;width:80pt;height:25pt"><x:ClientData ObjectType="Drop"><x:SizeWithCells/><x:Anchor>5, 0, 3, 0, 7, 0, 4, 0</x:Anchor><x:FmlaLink>$E$1</x:FmlaLink><x:FmlaRange>$F$1:$F$3</x:FmlaRange></x:ClientData></v:shape></xml>`
      )
    );
    edit("xl/worksheets/_rels/sheet1.xml.rels", text =>
      text.replace(
        "</Relationships>",
        `<Relationship Id="rIdDrop" Type="${CTRL_PROP}" Target="../ctrlProps/ctrlProp9.xml"/></Relationships>`
      )
    );
    edit("xl/worksheets/sheet1.xml", text =>
      text.replace(
        "</controls>",
        `<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="1030" r:id="rIdDrop" name="Drop Down 2"><controlPr defaultSize="0"><anchor moveWithCells="1"><from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls>`
      )
    );
    edit("[Content_Types].xml", text =>
      text.replace(
        "</Types>",
        `<Override PartName="/xl/ctrlProps/ctrlProp9.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>`
      )
    );
    return createZip([...files].map(([name, data]) => ({ name, data })));
  }

  it("moves the control in every place its anchor is written, and its cell links", async () => {
    const wb = await Workbook.read(await fixture());
    const ws = Workbook.getWorksheet(wb, "S")!;
    Row.insert(ws, 1, [], {});
    const files = await extractAll(await Workbook.toBuffer(wb));
    const text = (path: string) => new TextDecoder().decode(files.get(path)!.data);
    const vml = text("xl/drawings/vmlDrawing1.vml");
    expect(vml).toContain("<x:Anchor>5, 0, 4, 0, 7, 0, 5, 0</x:Anchor>");
    expect(vml).toContain("<x:FmlaLink>$E$2</x:FmlaLink>");
    expect(vml).toContain("<x:FmlaRange>$F$2:$F$4</x:FmlaRange>");
    const sheetXml = text("xl/worksheets/sheet1.xml");
    expect(sheetXml).toMatch(
      /name="Drop Down 2">.*<from><xdr:col>5<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>4<\/xdr:row>/
    );
    const props = [...files.keys()]
      .filter(name => name.startsWith("xl/ctrlProps/"))
      .map(name => text(name))
      .find(xml => xml.includes('objectType="Drop"'))!;
    expect(props).toContain('fmlaLink="$E$2"');
    expect(props).toContain('fmlaRange="$F$2:$F$4"');
  });
});

describe("hyperlink cells with a formula", () => {
  it("move with their formula, rewritten", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "B2", {
      formula: 'HYPERLINK("#S!A1",A2)',
      result: "go",
      hyperlink: "#S!A1"
    });
    Row.insert(ws, 1, []);
    Column.splice(ws, 1, 0, []);
    expect(Cell.getModel(ws, "C3")).toMatchObject({
      formula: 'HYPERLINK("#S!A1",B3)',
      hyperlink: "#S!A1"
    });
  });

  it("are rewritten on another sheet too", () => {
    const wb = Workbook.create();
    const data = Workbook.addWorksheet(wb, "Data");
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", { formula: "HYPERLINK(Data!A5)", result: "x", hyperlink: "#Data!A5" });
    Row.insert(data, 1, []);
    expect(Cell.getModel(ws, "A1").formula).toBe("HYPERLINK(Data!A6)");
  });
});

describe("structured references to a table the edit removes", () => {
  function build() {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const other = Workbook.addWorksheet(wb, "Other");
    Table.add(ws, { name: "T", ref: "A1", columns: [{ name: "x" }], rows: [[1], [2]] });
    Table.add(ws, { name: "Kept", ref: "C1", columns: [{ name: "y" }], rows: [[1]] });
    Cell.setValue(other, "A1", { formula: "SUM(T[x])+SUM(Kept[y])" });
    Cell.setValue(other, "A2", { formula: 'COUNTA(t[[#All],[x]])&"T[x]"' });
    DefinedNames.addFormula(Workbook.getDefinedNames(wb), "Total", "SUM(T[x])");
    ConditionalFormatting.add(other, {
      ref: "A1",
      rules: [{ type: "expression", formulae: ["A1>MAX(T[x])"], style: { font: { bold: true } } }]
    });
    return { wb, ws, other };
  }

  it("become #REF! when the table's rows are deleted", () => {
    const { wb, ws, other } = build();
    Column.splice(ws, 1, 1);
    expect(Table.get(ws, "T")).toBeUndefined();
    expect(Cell.getFormula(other, "A1")).toBe("SUM(#REF!)+SUM(Kept[y])");
    expect(Cell.getFormula(other, "A2")).toBe('COUNTA(#REF!)&"T[x]"');
    expect(DefinedNames.list(Workbook.getDefinedNames(wb))).toContainEqual(
      expect.objectContaining({ name: "Total", formulaExpression: "SUM(#REF!)" })
    );
    expect(ConditionalFormatting.list(other)[0].rules[0]).toMatchObject({
      formulae: ["A1>MAX(#REF!)"]
    });
  });

  it("become #REF! when the table's sheet is removed", () => {
    const { wb, other } = build();
    Workbook.removeWorksheet(wb, "S");
    expect(Cell.getFormula(other, "A1")).toBe("SUM(#REF!)+SUM(#REF!)");
  });
});
