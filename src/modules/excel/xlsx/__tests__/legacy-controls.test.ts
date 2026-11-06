/**
 * Legacy form controls other than checkboxes survive a rewrite.
 *
 * A button, a list or combo box, an option button, a group box, a spinner, a scroll bar and a label are
 * each a VML shape, a `<control>` entry in the sheet and an `xl/ctrlProps/ctrlPropN.xml`, exactly like a
 * checkbox — but only the checkbox is modelled. The writer regenerated the sheet's VML and `<controls>`
 * from the checkboxes alone, so every other control vanished from the drawing while its ctrlProp part and
 * relationship stayed behind with nothing pointing at them.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expectValidXlsb } from "@excel/__tests__/helpers/expect-valid-xlsb";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import { Cell, Form, Workbook, Worksheet } from "@excel/index";
import { describe, expect, it } from "vitest";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const CTRL_PROP_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

interface ControlSpec {
  id: number;
  objectType: string;
  name: string;
  /** `<x:ClientData>` children beyond the anchor. */
  clientData: string;
  /** `formControlPr` attributes. */
  props: string;
}

/** One of every control Excel's Forms toolbar draws, besides the checkbox. */
const CONTROLS: ControlSpec[] = [
  {
    id: 1030,
    objectType: "Button",
    name: "Button 6",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill><x:FmlaMacro>[0]!Macro1</x:FmlaMacro><x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>",
    props: 'objectType="Button" lockText="1"'
  },
  {
    id: 1031,
    objectType: "Drop",
    name: "Drop Down 7",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:AutoLine>False</x:AutoLine><x:FmlaLink>$E$1</x:FmlaLink><x:FmlaRange>$F$1:$F$3</x:FmlaRange><x:Val>0</x:Val><x:Sel>1</x:Sel><x:DropStyle>Combo</x:DropStyle><x:DropLines>8</x:DropLines>",
    props:
      'objectType="Drop" dropStyle="combo" dx="16" fmlaLink="$E$1" fmlaRange="$F$1:$F$3" sel="1"'
  },
  {
    id: 1032,
    objectType: "List",
    name: "List Box 8",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:FmlaRange>$F$1:$F$3</x:FmlaRange><x:Val>0</x:Val><x:Sel>0</x:Sel>",
    props: 'objectType="List" dx="16" fmlaRange="$F$1:$F$3" sel="0"'
  },
  {
    id: 1033,
    objectType: "Radio",
    name: "Option Button 9",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:FmlaLink>$E$2</x:FmlaLink><x:Checked>1</x:Checked><x:FirstButton/>",
    props: 'objectType="Radio" checked="Checked" firstButton="1" fmlaLink="$E$2" lockText="1"'
  },
  {
    id: 1034,
    objectType: "GBox",
    name: "Group Box 10",
    clientData: "<x:PrintObject>False</x:PrintObject><x:NoThreeD/>",
    props: 'objectType="GBox" noThreeD="1"'
  },
  {
    id: 1035,
    objectType: "Spin",
    name: "Spinner 11",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:FmlaLink>$E$3</x:FmlaLink><x:Val>3</x:Val><x:Min>0</x:Min><x:Max>10</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page>",
    props: 'objectType="Spin" dx="16" fmlaLink="$E$3" max="10" page="10" val="3"'
  },
  {
    id: 1036,
    objectType: "Scroll",
    name: "Scroll Bar 12",
    clientData:
      "<x:PrintObject>False</x:PrintObject><x:FmlaLink>$E$4</x:FmlaLink><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>100</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page><x:Horiz/>",
    props: 'objectType="Scroll" dx="16" fmlaLink="$E$4" horiz="1" max="100" page="10" val="0"'
  },
  {
    id: 1037,
    objectType: "Label",
    name: "Label 13",
    clientData: "<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill>",
    props: 'objectType="Label" lockText="1"'
  }
];

/** The second sheet's button reuses id 1030, which the first sheet's block already holds. */
const SECOND_SHEET_BUTTON: ControlSpec = { ...CONTROLS[0], name: "Button 1" };

const SHAPETYPE_201 = `<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>`;

function vmlShape(spec: ControlSpec, row: number): string {
  return `<v:shape id="${spec.name.replace(/ /g, "_x0020_")}" o:spid="_x0000_s${spec.id}" type="#_x0000_t201" style="position:absolute;margin-left:200pt;margin-top:${row * 20}pt;width:80pt;height:25pt;z-index:${row};mso-wrap-style:tight" o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto"><v:fill color2="buttonFace [67]" o:detectmouseclick="t"/><o:lock v:ext="edit" rotation="t"/><v:textbox style="mso-direction-alt:auto" o:singleclick="f"><div style="text-align:center"><font face="Calibri" size="220" color="#000000">${spec.objectType} caption</font></div></v:textbox><x:ClientData ObjectType="${spec.objectType}"><x:Anchor>5, 0, ${row}, 0, 7, 0, ${row + 1}, 0</x:Anchor>${spec.clientData}</x:ClientData></v:shape>`;
}

function controlEntry(spec: ControlSpec, rId: string, row: number): string {
  return `<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="${spec.id}" r:id="${rId}" name="${spec.name}"><controlPr defaultSize="0" print="0" autoFill="0" autoPict="0"${spec.objectType === "Button" ? ' macro="[0]!Macro1"' : ""}><anchor moveWithCells="1"><from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent>`;
}

function ctrlPropXml(spec: ControlSpec): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<formControlPr xmlns="${CTRL_PROP_NS}" ${spec.props}/>`;
}

const CTRL_PROP_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const VML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

/**
 * Sheet `S`: a note, a checkbox this library wrote, and every other control Excel draws, numbered with
 * gaps (`ctrlProp10`…) the way a workbook edited in Excel ends up. Sheet `T`: a lone button in a VML
 * part of its own, with no checkbox and no note beside it.
 */
async function fixture(controls: readonly ControlSpec[] = CONTROLS): Promise<Uint8Array> {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Workbook.addWorksheet(wb, "T");
  Cell.setValue(ws, "A1", 1);
  Cell.setNote(ws, "A1", "hello");
  Form.addCheckbox(ws, "B2", { checked: true, link: "C2", text: "cb" });

  const files = new Map<string, Uint8Array>();
  for (const [name, entry] of await extractAll(await Workbook.toBuffer(wb))) {
    files.set(name, entry.data);
  }
  const edit = (path: string, change: (text: string) => string) =>
    files.set(path, encoder.encode(change(decoder.decode(files.get(path)!))));

  const sheetRels: string[] = [];
  const sheetControls: string[] = [];
  const shapes: string[] = [];
  const overrides: string[] = [];
  controls.forEach((spec, i) => {
    const part = `ctrlProp${10 + i}.xml`;
    const rId = `rIdCtl${i}`;
    files.set(`xl/ctrlProps/${part}`, encoder.encode(ctrlPropXml(spec)));
    sheetRels.push(
      `<Relationship Id="${rId}" Type="${CTRL_PROP_REL}" Target="../ctrlProps/${part}"/>`
    );
    sheetControls.push(controlEntry(spec, rId, 3 + i * 2));
    shapes.push(vmlShape(spec, 3 + i * 2));
    overrides.push(
      `<Override PartName="/xl/ctrlProps/${part}" ContentType="application/vnd.ms-excel.controlproperties+xml"/>`
    );
  });
  edit("xl/drawings/vmlDrawing1.vml", text => text.replace("</xml>", `${shapes.join("")}</xml>`));
  edit("xl/worksheets/_rels/sheet1.xml.rels", text =>
    text.replace("</Relationships>", `${sheetRels.join("")}</Relationships>`)
  );
  edit("xl/worksheets/sheet1.xml", text =>
    text.replace("</controls>", `${sheetControls.join("")}</controls>`)
  );

  // Sheet T.
  files.set("xl/ctrlProps/ctrlProp30.xml", encoder.encode(ctrlPropXml(SECOND_SHEET_BUTTON)));
  overrides.push(
    `<Override PartName="/xl/ctrlProps/ctrlProp30.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>`
  );
  files.set(
    "xl/drawings/vmlDrawing9.vml",
    encoder.encode(
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>${SHAPETYPE_201}${vmlShape(SECOND_SHEET_BUTTON, 1)}</xml>`
    )
  );
  files.set(
    "xl/worksheets/_rels/sheet2.xml.rels",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${VML_REL}" Target="../drawings/vmlDrawing9.vml"/><Relationship Id="rId2" Type="${CTRL_PROP_REL}" Target="../ctrlProps/ctrlProp30.xml"/></Relationships>`
    )
  );
  edit("xl/worksheets/sheet2.xml", text =>
    text
      .replace(
        "<worksheet ",
        '<worksheet xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
      )
      .replace(
        "</worksheet>",
        `<legacyDrawing r:id="rId1"/><mc:AlternateContent><mc:Choice Requires="x14"><controls>${controlEntry(SECOND_SHEET_BUTTON, "rId2", 1)}</controls></mc:Choice></mc:AlternateContent></worksheet>`
      )
  );
  edit("[Content_Types].xml", text => text.replace("</Types>", `${overrides.join("")}</Types>`));
  return createZip([...files].map(([name, data]) => ({ name, data })));
}

async function readBack(bytes: Uint8Array, format?: "xlsb") {
  const wb = Workbook.create();
  await Workbook.read(wb, bytes, format ? { format } : undefined);
  return wb;
}

type Parts = Map<string, string>;

async function parts(bytes: Uint8Array): Promise<Parts> {
  const out = new Map<string, string>();
  for (const [name, entry] of await extractAll(bytes)) {
    out.set(name, decoder.decode(entry.data));
  }
  return out;
}

function attr(tag: string, name: string): string | undefined {
  return new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];
}

/** Each `<v:shape>` of a VML part: its ObjectType, its shape id, and the whole element. */
function vmlShapes(vml: string): { objectType: string; id: number; xml: string }[] {
  return [...vml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map(([xml]) => {
    const open = /^<v:shape\b[^>]*>/.exec(xml)![0];
    const spid = attr(open, "o:spid") ?? attr(open, "id")!;
    return {
      objectType: /ObjectType="([^"]*)"/.exec(xml)?.[1] ?? "",
      id: Number(/_x0000_s(\d+)/.exec(spid)![1]),
      xml
    };
  });
}

/** Each `<control>` of a sheet. */
function sheetControls(
  sheet: string
): { shapeId: number; rId: string; name: string; xml: string }[] {
  return [...sheet.matchAll(/<control\b[^>]*>[\s\S]*?<\/control>/g)].map(([xml]) => {
    const open = /^<control\b[^>]*>/.exec(xml)![0];
    return {
      shapeId: Number(attr(open, "shapeId")),
      rId: attr(open, "r:id")!,
      name: attr(open, "name")!,
      xml
    };
  });
}

function rels(xml: string | undefined): { id: string; type: string; target: string }[] {
  return [...(xml ?? "").matchAll(/<Relationship\b[^>]*>/g)].map(([tag]) => ({
    id: attr(tag, "Id")!,
    type: attr(tag, "Type")!,
    target: attr(tag, "Target")!
  }));
}

/** A sheet's controls, joined through their relationships to the ctrlProp each one reaches. */
function controlsOf(files: Parts, sheetIndex: number) {
  const sheetRels = rels(files.get(`xl/worksheets/_rels/sheet${sheetIndex}.xml.rels`));
  return sheetControls(files.get(`xl/worksheets/sheet${sheetIndex}.xml`)!).map(control => {
    const rel = sheetRels.find(r => r.id === control.rId);
    const path = rel ? `xl/${rel.target.replace(/^\.\.\//, "")}` : undefined;
    return { ...control, relType: rel?.type, ctrlProp: path && files.get(path) };
  });
}

function vmlOf(files: Parts, sheetIndex: number, ext = "xml"): string {
  const vmlRel = rels(files.get(`xl/worksheets/_rels/sheet${sheetIndex}.${ext}.rels`)).find(
    rel => rel.type === VML_REL
  )!;
  return files.get(`xl/${vmlRel.target.replace(/^\.\.\//, "")}`)!;
}

/** The shape with its id attributes removed, so a renumbered shape compares equal to its source. */
function withoutIds(xml: string): string {
  return xml.replace(/\s(?:id|o:spid)="[^"]*"/g, "");
}

describe("legacy form controls other than checkboxes", () => {
  it("keeps every control's shape, entry and properties through a rewrite", async () => {
    const source = await parts(await fixture());
    const out = await Workbook.toBuffer(await readBack(await fixture()));
    await expectValidXlsx(out);
    const files = await parts(out);

    const controls = controlsOf(files, 1);
    expect(controls.map(c => c.name)).toEqual(["Check Box 1", ...CONTROLS.map(c => c.name)]);
    const shapes = vmlShapes(vmlOf(files, 1));
    const sourceShapes = vmlShapes(vmlOf(source, 1));
    for (const spec of CONTROLS) {
      const control = controls.find(c => c.name === spec.name)!;
      expect(control.relType).toBe(CTRL_PROP_REL);
      // The part's content is the one read, verbatim.
      expect(control.ctrlProp).toBe(ctrlPropXml(spec));
      // The `<control>` says what it said, and names the shape that draws it.
      const sourceControl = sheetControls(source.get("xl/worksheets/sheet1.xml")!).find(
        c => c.name === spec.name
      )!;
      expect(control.xml.replace(/\s(?:shapeId|r:id)="[^"]*"/g, "")).toBe(
        sourceControl.xml.replace(/\s(?:shapeId|r:id)="[^"]*"/g, "")
      );
      const shape = shapes.find(s => s.id === control.shapeId)!;
      expect(shape.objectType).toBe(spec.objectType);
      expect(withoutIds(shape.xml)).toBe(
        withoutIds(sourceShapes.find(s => s.objectType === spec.objectType)!.xml)
      );
    }
    // The note and the checkbox are still there.
    expect(shapes.filter(s => s.objectType === "Note")).toHaveLength(1);
    expect(shapes.filter(s => s.objectType === "Checkbox")).toHaveLength(1);
    expect(files.get("xl/comments1.xml")).toContain("hello");
  });

  it("gives the second sheet's button a free id and keeps its own shapetype", async () => {
    const files = await parts(await Workbook.toBuffer(await readBack(await fixture())));
    const [button] = controlsOf(files, 2);
    expect(button.name).toBe("Button 1");
    expect(button.ctrlProp).toBe(ctrlPropXml(SECOND_SHEET_BUTTON));
    const vml = vmlOf(files, 2);
    expect(vml).toContain('id="_x0000_t201"');
    const [shape] = vmlShapes(vml);
    expect(shape.objectType).toBe("Button");
    expect(shape.id).toBe(button.shapeId);
    // Shape ids are unique across the workbook.
    const all = [...vmlShapes(vmlOf(files, 1)), ...vmlShapes(vml)].map(s => s.id);
    expect(new Set(all).size).toBe(all.length);
    expect(attr(/<o:idmap[^>]*>/.exec(vml)![0], "data")).toBe(
      String(Math.floor(button.shapeId / 1024))
    );
  });

  it("leaves no orphan part and no dangling relationship", async () => {
    const files = await parts(await Workbook.toBuffer(await readBack(await fixture())));
    const ctrlPropParts = [...files.keys()].filter(p => p.startsWith("xl/ctrlProps/"));
    const reached = [1, 2].flatMap(i =>
      rels(files.get(`xl/worksheets/_rels/sheet${i}.xml.rels`))
        .filter(rel => rel.type === CTRL_PROP_REL)
        .map(rel => `xl/${rel.target.replace(/^\.\.\//, "")}`)
    );
    expect(ctrlPropParts.sort()).toEqual([...reached].sort());
    expect(ctrlPropParts).toHaveLength(1 + CONTROLS.length + 1);
    // Every relationship is named by a `<control>`, and every one reaches a part.
    for (const i of [1, 2]) {
      const named = new Set(controlsOf(files, i).map(c => c.rId));
      for (const rel of rels(files.get(`xl/worksheets/_rels/sheet${i}.xml.rels`))) {
        if (rel.type === CTRL_PROP_REL) {
          expect(named.has(rel.id)).toBe(true);
        }
      }
    }
    const types = files.get("[Content_Types].xml")!;
    for (const part of ctrlPropParts) {
      expect(types).toContain(`PartName="/${part}"`);
    }
  });

  it("is stable over a second round trip, and checkboxes stay editable beside the rest", async () => {
    const first = await readBack(await fixture());
    const ws = Workbook.getWorksheet(first, "S")!;
    expect(Form.list(ws)).toHaveLength(1);
    Form.remove(ws, Form.list(ws)[0]);
    Form.addCheckbox(ws, "B20", { text: "new", link: "c9" });
    const once = await Workbook.toBuffer(first);
    const twice = await Workbook.toBuffer(await readBack(once));
    await expectValidXlsx(twice);
    const a = await parts(once);
    const b = await parts(twice);
    for (const i of [1, 2]) {
      expect(controlsOf(b, i).map(c => [c.name, c.ctrlProp])).toEqual(
        controlsOf(a, i).map(c => [c.name, c.ctrlProp])
      );
      expect(vmlShapes(vmlOf(b, i)).map(s => s.objectType)).toEqual(
        vmlShapes(vmlOf(a, i)).map(s => s.objectType)
      );
    }
    const controls = controlsOf(b, 1);
    expect(controls.filter(c => c.ctrlProp?.includes('objectType="CheckBox"'))).toHaveLength(1);
    // The new checkbox did not take an id a preserved control holds.
    const ids = controls.map(c => c.shapeId);
    expect(new Set(ids).size).toBe(ids.length);
    const [read] = Form.list(Workbook.getWorksheet(await readBack(twice), "S")!);
    expect(Form.getLink(read)).toBe("$C$9");
  });

  it("carries the controls through a copied worksheet model", async () => {
    const wb = await readBack(await fixture());
    const target = Workbook.create();
    const copy = Workbook.addWorksheet(target, "Copy");
    Worksheet.setModel(copy, Worksheet.getModel(Workbook.getWorksheet(wb, "T")!));
    const files = await parts(await Workbook.toBuffer(target));
    const [button] = controlsOf(files, 1);
    expect(button?.name).toBe("Button 1");
    expect(vmlShapes(vmlOf(files, 1)).map(s => s.objectType)).toEqual(["Button"]);
  });
});

describe("legacy form controls through an XLSB", () => {
  /** Each sheet's ctrlProp parts, by content, and whether every one is reached from its sheet. */
  function ctrlPropsOf(files: Parts, sheetIndex: number): string[] {
    return rels(files.get(`xl/worksheets/_rels/sheet${sheetIndex}.bin.rels`))
      .filter(rel => rel.type === CTRL_PROP_REL)
      .map(rel => files.get(`xl/${rel.target.replace(/^\.\.\//, "")}`)!);
  }

  it("keeps every control's shape and properties part, with nothing orphaned", async () => {
    const xlsb = await Workbook.toBuffer(await readBack(await fixture()), { format: "xlsb" });
    await expectValidXlsb(xlsb);
    const files = await parts(xlsb);
    const types = vmlShapes(vmlOf(files, 1, "bin")).map(s => s.objectType);
    expect(types).toEqual(["Note", "Checkbox", ...CONTROLS.map(c => c.objectType)]);
    expect(vmlShapes(vmlOf(files, 2, "bin")).map(s => s.objectType)).toEqual(["Button"]);
    const props = ctrlPropsOf(files, 1);
    expect(props.slice(1)).toEqual(CONTROLS.map(ctrlPropXml));
    expect(ctrlPropsOf(files, 2)).toEqual([ctrlPropXml(SECOND_SHEET_BUTTON)]);
    const written = [...files.keys()].filter(p => p.startsWith("xl/ctrlProps/"));
    expect(written).toHaveLength(props.length + 1);
  });

  it("round-trips an XLSB with them, and converts back to XLSX keeping the shapes", async () => {
    const first = await readBack(
      await Workbook.toBuffer(await readBack(await fixture()), { format: "xlsb" }),
      "xlsb"
    );
    expect(Form.list(Workbook.getWorksheet(first, "S")!)).toHaveLength(1);
    // An edit, so the package is written rather than handed back as read.
    Cell.setValue(Workbook.getWorksheet(first, "S")!, "Z1", 1);
    const again = await Workbook.toBuffer(first, { format: "xlsb" });
    await expectValidXlsb(again);
    const files = await parts(again);
    expect(vmlShapes(vmlOf(files, 1, "bin")).map(s => s.objectType)).toEqual([
      "Note",
      "Checkbox",
      ...CONTROLS.map(c => c.objectType)
    ]);
    expect(ctrlPropsOf(files, 1).slice(1)).toEqual(CONTROLS.map(ctrlPropXml));

    // An XLSB names no control in a `<controls>` entry, so the XLSX keeps each one as Excel 2007 stored it:
    // the VML shape, which carries the control whole. No ctrlProp part is written that nothing would reach.
    const xlsx = await Workbook.toBuffer(first);
    await expectValidXlsx(xlsx);
    const converted = await parts(xlsx);
    expect(vmlShapes(vmlOf(converted, 1)).map(s => s.objectType)).toEqual([
      "Note",
      "Checkbox",
      ...CONTROLS.map(c => c.objectType)
    ]);
    expect([...converted.keys()].filter(p => p.startsWith("xl/ctrlProps/"))).toHaveLength(1);
  });
});

/**
 * Two controls of one type whose properties differ. An XLSB carries no key from a control to its ctrlProp
 * part — [MS-XLSB] gives the worksheet no relationship to one and no record naming one — so the reader
 * pairs each part with the shape whose own `<x:ClientData>` says the same thing.
 */
describe("same-type controls through an XLSB", () => {
  const listBox = (id: number, name: string, range: string, sel: number): ControlSpec => ({
    id,
    objectType: "List",
    name,
    clientData: `<x:FmlaRange>${range}</x:FmlaRange><x:Val>0</x:Val><x:Sel>${sel}</x:Sel>`,
    props: `objectType="List" dx="16" fmlaRange="${range}" sel="${sel}"`
  });
  const LISTS = [
    listBox(1030, "List Box 6", "$F$1:$F$3", 1),
    listBox(1031, "List Box 7", "$G$1:$G$9", 2)
  ];

  /** The XLSB written from the fixture, with its two ctrlProp relationships listed in reverse order. */
  async function reversedXlsb(): Promise<Uint8Array> {
    const xlsb = await Workbook.toBuffer(await readBack(await fixture(LISTS)), { format: "xlsb" });
    const files = new Map<string, Uint8Array>();
    for (const [name, entry] of await extractAll(xlsb)) {
      files.set(name, entry.data);
    }
    const relsPath = "xl/worksheets/_rels/sheet1.bin.rels";
    const relsXml = decoder.decode(files.get(relsPath)!);
    const listRels = [...relsXml.matchAll(/<Relationship\b[^>]*>/g)]
      .map(([tag]) => tag)
      .filter(tag => tag.includes(CTRL_PROP_REL))
      .filter(tag =>
        decoder
          .decode(files.get(`xl/${attr(tag, "Target")!.replace(/^\.\.\//, "")}`)!)
          .includes('objectType="List"')
      );
    expect(listRels).toHaveLength(2);
    const [first, second] = listRels;
    const swapped = relsXml
      .replace(first, "\u0000")
      .replace(second, first)
      .replace("\u0000", second);
    files.set(relsPath, encoder.encode(swapped));
    return createZip([...files].map(([name, data]) => ({ name, data })));
  }

  /** Each list box's own range in its VML, with the range its kept ctrlProp says. */
  function pairs(ws: Worksheet.Handle): [string | undefined, string | undefined][] {
    return (Worksheet.getModel(ws).legacyControls ?? [])
      .filter(control => control.objectType === "List")
      .map(control => {
        const clientData = control.vmlShape.children.find(child => child.tag === "x:ClientData");
        const range = clientData?.children.find(child => child.tag === "x:FmlaRange");
        const text = (range?.content ?? []).filter(item => typeof item === "string").join("");
        return [text, /fmlaRange="([^"]*)"/.exec(control.ctrlPropXml ?? "")?.[1]];
      });
  }

  it("pairs each ctrlProp with its own shape, whatever order the relationships are in", async () => {
    const wb = await readBack(await reversedXlsb(), "xlsb");
    expect(pairs(Workbook.getWorksheet(wb, "S")!)).toEqual([
      ["$F$1:$F$3", "$F$1:$F$3"],
      ["$G$1:$G$9", "$G$1:$G$9"]
    ]);
  });

  it("keeps each control's own properties through xlsb→xlsb and xlsb→xlsx", async () => {
    const wb = await readBack(await reversedXlsb(), "xlsb");
    Cell.setValue(Workbook.getWorksheet(wb, "S")!, "Z1", 1);
    const again = await readBack(await Workbook.toBuffer(wb, { format: "xlsb" }), "xlsb");
    expect(pairs(Workbook.getWorksheet(again, "S")!)).toEqual([
      ["$F$1:$F$3", "$F$1:$F$3"],
      ["$G$1:$G$9", "$G$1:$G$9"]
    ]);
    const xlsx = await parts(await Workbook.toBuffer(again));
    const lists = vmlShapes(vmlOf(xlsx, 1)).filter(shape => shape.objectType === "List");
    expect(lists.map(shape => /<x:FmlaRange>([^<]*)</.exec(shape.xml)?.[1])).toEqual([
      "$F$1:$F$3",
      "$G$1:$G$9"
    ]);
    expect(lists.map(shape => /<x:Sel>([^<]*)</.exec(shape.xml)?.[1])).toEqual(["1", "2"]);
  });

  it("falls back to relationship order only between controls their properties cannot tell apart", async () => {
    const twins = [
      listBox(1030, "List Box 6", "$F$1:$F$3", 1),
      listBox(1031, "List Box 7", "$F$1:$F$3", 1)
    ];
    const wb = await readBack(
      await Workbook.toBuffer(await readBack(await fixture(twins)), { format: "xlsb" }),
      "xlsb"
    );
    expect(pairs(Workbook.getWorksheet(wb, "S")!)).toEqual([
      ["$F$1:$F$3", "$F$1:$F$3"],
      ["$F$1:$F$3", "$F$1:$F$3"]
    ]);
  });
});
