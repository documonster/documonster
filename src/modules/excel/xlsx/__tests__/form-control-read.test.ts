/**
 * Legacy form checkboxes survive a read.
 *
 * A checkbox is three parts — a VML shape (`<x:ClientData ObjectType="Checkbox">`), an
 * `xl/ctrlProps/ctrlPropN.xml`, and a `<controls>` entry in the sheet bridging the two — and the reader
 * understood none of them, so `Form.list` came back empty for every file holding one, including files this
 * library wrote, and the next save dropped them.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsb } from "@excel/__tests__/helpers/expect-valid-xlsb";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Cell, Form, Image, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function libraryWorkbook() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", true);
  Form.addCheckbox(ws, "B2", { text: "Agree", checked: true, link: "A1" });
  Form.addCheckbox(ws, "B5:D6", { text: "Second", noThreeD: false, print: true });
  return wb;
}

async function readBack(bytes: Uint8Array, format?: "xlsb") {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes, format ? ({ format } as never) : undefined);
  return wb;
}

function summary(wb: ReturnType<typeof Workbook.create>) {
  const ws = Workbook.getWorksheet(wb, "S")!;
  return Form.list(ws).map(cb => ({
    checked: Form.isChecked(cb),
    link: Form.getLink(cb),
    text: Form.getText(cb),
    shapeId: cb.model.shapeId,
    name: cb.model.name,
    tl: cb.model.tl,
    br: cb.model.br,
    noThreeD: cb.model.noThreeD,
    print: cb.model.print
  }));
}

async function partText(bytes: Uint8Array, path: string): Promise<string | undefined> {
  const entry = (await extractAll(bytes)).get(path);
  return entry ? decoder.decode(entry.data) : undefined;
}

describe.each([["xlsx"], ["xlsb"]] as const)("library-written checkboxes (%s)", format => {
  const options = format === "xlsb" ? ({ format } as const) : undefined;
  const validate = format === "xlsb" ? expectValidXlsb : expectValidXlsx;

  it("reads every checkbox with its state, link, text and anchor", async () => {
    const source = libraryWorkbook();
    // An XLSX names each control in `<controls>`. An XLSB has no such element, and this library's VML
    // shapes carry only their id, so the name stays absent — which a write turns into the same default.
    const expected = summary(source).map(c => ({
      ...c,
      name: format === "xlsb" ? undefined : c.shapeId === 1025 ? "Check Box 1" : "Check Box 2"
    }));
    const wb = await readBack(await Workbook.toBuffer(source, options), options?.format);
    expect(summary(wb)).toEqual(expected);
  });

  it("round-trips them: same controls, stable names, one set of parts", async () => {
    const first = await readBack(
      await Workbook.toBuffer(libraryWorkbook(), options),
      options?.format
    );
    // An edit, so an XLSB is written rather than handed back as the bytes it was read from.
    Cell.setValue(Workbook.getWorksheet(first, "S")!, "Z1", 1);
    const out = await Workbook.toBuffer(first, options);
    await validate(out);
    const parts = [...(await extractAll(out)).keys()];
    expect(parts.filter(p => p.startsWith("xl/ctrlProps/")).sort()).toEqual([
      "xl/ctrlProps/ctrlProp1.xml",
      "xl/ctrlProps/ctrlProp2.xml"
    ]);
    expect(parts.filter(p => /vmlDrawing\d+\.vml$/.test(p))).toHaveLength(1);
    const vml = (await partText(out, "xl/drawings/vmlDrawing1.vml"))!;
    expect(vml.match(/ObjectType="Checkbox"/g)).toHaveLength(2);
    // The hidden DrawingML bridge anchors are regenerated, not kept beside the new ones.
    const drawing = (await partText(out, "xl/drawings/drawing1.xml"))!;
    expect(drawing.match(/_x0000_s1025/g)).toHaveLength(1);
    expect(drawing.match(/_x0000_s1026/g)).toHaveLength(1);
    expect(drawing).toContain('name="Check Box 1"');
    expect(drawing).toContain('name="Check Box 2"');
    const second = await readBack(out, options?.format);
    expect(summary(second)).toEqual(summary(first));
  });

  it("lets a read checkbox be edited and removed", async () => {
    const wb = await readBack(await Workbook.toBuffer(libraryWorkbook(), options), options?.format);
    const ws = Workbook.getWorksheet(wb, "S")!;
    const [first, second] = Form.list(ws);
    Form.setChecked(first, false);
    Form.remove(ws, second);
    const again = await readBack(await Workbook.toBuffer(wb, options), options?.format);
    const list = Form.list(Workbook.getWorksheet(again, "S")!);
    expect(list).toHaveLength(1);
    expect(Form.isChecked(list[0])).toBe(false);
    expect(Form.getText(list[0])).toBe("Agree");
  });
});

describe.each([["xlsx"], ["xlsb"]] as const)("checkbox link addresses (%s)", format => {
  const options = format === "xlsb" ? ({ format } as const) : undefined;
  const cases: [string, string][] = [
    ["c5", "$C$5"],
    ["C5", "$C$5"],
    ["$C5", "$C$5"],
    ["$C$5", "$C$5"],
    ["Other!b2", "Other!$B$2"],
    ["Other!$B$2", "Other!$B$2"],
    ["'My Sheet'!A1", "'My Sheet'!$A$1"],
    ["'It''s'!a1:b2", "'It''s'!$A$1:$B$2"],
    ["My Sheet!A1", "'My Sheet'!$A$1"],
    ["=D4", "$D$4"],
    ["a1:b2", "$A$1:$B$2"],
    ["MyFlag", "MyFlag"],
    ["Other!MyFlag", "Other!MyFlag"]
  ];

  it("normalises each to an absolute reference and keeps the sheet", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "A");
    Workbook.addWorksheet(wb, "Other");
    Workbook.addWorksheet(wb, "My Sheet");
    Workbook.addWorksheet(wb, "It's");
    cases.forEach(([link], i) => Form.addCheckbox(ws, `B${i + 2}`, { link }));
    const expected = cases.map(([, absolute]) => absolute);
    expect(Form.list(ws).map(cb => Form.getLink(cb))).toEqual(expected);

    const out = await Workbook.toBuffer(wb, options);
    const read = await readBack(out, options?.format);
    expect(Form.list(Workbook.getWorksheet(read, "A")!).map(cb => Form.getLink(cb))).toEqual(
      expected
    );
    if (format === "xlsx") {
      const prop = (await partText(out, "xl/ctrlProps/ctrlProp1.xml"))!;
      expect(prop).toContain('fmlaLink="$C$5"');
      const vml = (await partText(out, "xl/drawings/vmlDrawing1.vml"))!;
      expect(vml).toMatch(/<x:FmlaLink>(?:'|&apos;)My Sheet(?:'|&apos;)!\$A\$1<\/x:FmlaLink>/);
    }
  });

  it("normalises a link set later the same way", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "A");
    const cb = Form.addCheckbox(ws, "B2");
    Form.setLink(cb, "'My Sheet'!c3");
    expect(Form.getLink(cb)).toBe("'My Sheet'!$C$3");
  });
});

const EXCEL_VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype><v:shape id="Check_x0020_Box_x0020_7" o:spid="_x0000_s1031" type="#_x0000_t201"
  style='position:absolute;margin-left:48pt;margin-top:15pt;width:96pt;height:18.75pt;z-index:1;mso-wrap-style:tight'
  filled="f" fillcolor="window [65]" stroked="f" strokecolor="windowText [64]"
  o:insetmode="auto">
  <v:path shadowok="t" strokeok="t" fillok="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:left'><font face="Segoe UI" size="160" color="auto">I &amp; agree</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:Anchor>
    1, 0, 1, 0, 3, 0, 2, 5</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:TextVAlign>Center</x:TextVAlign>
   <x:FmlaLink>$D$2</x:FmlaLink>
   <x:NoThreeD/>
   <x:Checked>1</x:Checked>
  </x:ClientData>
 </v:shape></xml>`;

const EXCEL_CTRL_PROP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked" fmlaLink="$D$2" lockText="1" noThreeD="1"/>`;

const EXCEL_CONTROLS = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><control shapeId="1031" r:id="rIdCtl" name="Check Box 7"><controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>47625</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent>`;

/**
 * A sheet shaped the way Excel writes one checkbox: its own VML (named shape, Excel's attributes, a
 * whitespace-padded anchor, `checked="Checked"` in the ctrlProp), a `<controls>` entry under
 * `mc:AlternateContent`, and no DrawingML bridge.
 */
async function excelShapedPackage(): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "D2", true);
  const files = new Map<string, Uint8Array>();
  for (const [name, entry] of await extractAll(await Workbook.toBuffer(wb))) {
    files.set(name, entry.data);
  }
  const edit = (path: string, change: (text: string) => string) =>
    files.set(path, encoder.encode(change(decoder.decode(files.get(path)!))));
  files.set("xl/drawings/vmlDrawing1.vml", encoder.encode(EXCEL_VML));
  files.set("xl/ctrlProps/ctrlProp1.xml", encoder.encode(EXCEL_CTRL_PROP));
  files.set(
    "xl/worksheets/_rels/sheet1.xml.rels",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rIdCtl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/></Relationships>`
    )
  );
  edit("[Content_Types].xml", text =>
    text.replace(
      "</Types>",
      `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>`
    )
  );
  edit("xl/worksheets/sheet1.xml", text => {
    let out = text;
    if (!out.includes('xmlns:xdr="')) {
      out = out.replace(
        "<worksheet ",
        '<worksheet xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
      );
    }
    if (!out.includes('xmlns:r="')) {
      out = out.replace(
        "<worksheet ",
        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      );
    }
    return out.replace(
      "</worksheet>",
      `<legacyDrawing r:id="rIdVml"/>${EXCEL_CONTROLS}</worksheet>`
    );
  });
  return createZip([...files].map(([name, data]) => ({ name, data })));
}

describe("an Excel-shaped checkbox", () => {
  it("reads the checkbox Excel wrote", async () => {
    const wb = await readBack(await excelShapedPackage());
    const [cb, ...rest] = Form.list(Workbook.getWorksheet(wb, "S")!);
    expect(rest).toEqual([]);
    expect(Form.isChecked(cb)).toBe(true);
    expect(Form.getLink(cb)).toBe("$D$2");
    expect(Form.getText(cb)).toBe("I & agree");
    expect(cb.model.shapeId).toBe(1031);
    expect(cb.model.name).toBe("Check Box 7");
    expect(cb.model.noThreeD).toBe(true);
    // The `<controls>` anchor is in EMU and exact; the VML one is in pixels.
    expect(cb.model.tl).toEqual({ col: 1, colOff: 0, row: 1, rowOff: 0 });
    expect(cb.model.br).toEqual({ col: 3, colOff: 0, row: 2, rowOff: 47625 });
    expect(cb.model.vmlStyle).toContain("height:18.75pt");
  });

  it("round-trips it as a valid file with the same control", async () => {
    const wb = await readBack(await excelShapedPackage());
    const before = summary(wb);
    const out = await Workbook.toBuffer(wb);
    await expectValidXlsx(out);
    const sheet = (await partText(out, "xl/worksheets/sheet1.xml"))!;
    expect(sheet.match(/<control /g)).toHaveLength(1);
    expect(sheet).toContain('name="Check Box 7"');
    const vml = (await partText(out, "xl/drawings/vmlDrawing1.vml"))!;
    expect(vml).toContain("height:18.75pt");
    expect(summary(await readBack(out))).toEqual(before);
  });

  it("falls back to the VML when the ctrlProp part is missing", async () => {
    const files = await extractAll(await excelShapedPackage());
    const kept = [...files]
      .filter(([name]) => name !== "xl/ctrlProps/ctrlProp1.xml")
      .map(([name, entry]) => ({ name, data: entry.data }));
    const wb = await readBack(await createZip(kept));
    const [cb] = Form.list(Workbook.getWorksheet(wb, "S")!);
    expect(Form.isChecked(cb)).toBe(true);
    expect(Form.getLink(cb)).toBe("$D$2");
    expect(cb.model.br).toEqual({ col: 3, colOff: 0, row: 2, rowOff: 47625 });
  });
});

describe("an XLSB sheet whose drawing also holds a picture", () => {
  const PNG =
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  it("keeps the preserved drawing and the checkbox, and reports no loss", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, id, "E2:F4");
    Form.addCheckbox(ws, "B2", { text: "Agree", checked: true });
    const read = await readBack(await Workbook.toBuffer(wb, { format: "xlsb" }), "xlsb");
    expect(Form.list(Workbook.getWorksheet(read, "S")!)).toHaveLength(1);
    // An edit, so the package is written rather than handed back as the bytes it was read from.
    Cell.setValue(Workbook.getWorksheet(read, "S")!, "Z1", 1);
    // Before the checkbox was read, the sheet carried no control and the preserved drawing was the only
    // one; reading it must not turn the next write into a refused "cannot merge" loss.
    const out = await Workbook.toBuffer(read, { format: "xlsb" });
    await expectValidXlsb(out);
    const parts = [...(await extractAll(out)).keys()];
    expect(parts.filter(p => /drawing\d+\.xml$/.test(p))).toHaveLength(1);
    expect(parts.some(p => p.startsWith("xl/media/"))).toBe(true);
    expect(summary(await readBack(out, "xlsb"))).toEqual(summary(read));
  });
});
