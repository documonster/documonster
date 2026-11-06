/**
 * A DrawingML shape read from a file keeps everything its fields do not
 * model — custom geometry, the style matrix, a gradient dashed line with
 * arrowheads, text-box insets and anchoring, extension lists — and writes it
 * back in schema order. Fields the shape does model stay authoritative: an
 * edited one replaces the markup it was read from.
 */
import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { parsePartXml } from "@word/reader/part-xml";
import type { XmlElement } from "@xml/types";
import { describe, expect, it } from "vitest";

import { Document, Io, Transform } from "../index";
import type { DocxDocument, DrawingShape } from "../types";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

/** The wsp subtree of the fixture, as Word writes one. */
const WSP =
  `<wps:wsp>` +
  `<wps:cNvPr id="1" name="Freeform 1" descr="A custom shape"/>` +
  `<wps:cNvSpPr><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>` +
  `<wps:spPr bwMode="auto">` +
  `<a:xfrm rot="600000"><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>` +
  `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
  `<a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
  `<a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst>` +
  `</a:custGeom>` +
  `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>` +
  `<a:ln w="28575" cap="rnd" cmpd="dbl" algn="ctr">` +
  `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>` +
  `<a:lin ang="0" scaled="1"/></a:gradFill>` +
  `<a:custDash><a:ds d="300000" sp="100000"/></a:custDash><a:round/>` +
  `<a:headEnd type="triangle" w="med" len="med"/><a:tailEnd type="stealth"/></a:ln>` +
  `<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">` +
  `<a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>` +
  `</wps:spPr>` +
  `<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>` +
  `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
  `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>` +
  `<wps:txbx id="7"><w:txbxContent><w:p><w:r><w:t>Inside</w:t></w:r></w:p></w:txbxContent>` +
  `<wps:extLst><a:ext uri="{00000000-0000-0000-0000-000000000001}"/></wps:extLst></wps:txbx>` +
  `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" ` +
  `anchor="ctr" anchorCtr="0" upright="1"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>` +
  `<a:noAutofit/></wps:bodyPr>` +
  `</wps:wsp>`;

const ANCHORED_SHAPE =
  `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
  `relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="1" name="Freeform 1" descr="A custom shape"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  WSP +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for await (const entry of unzip(bytes).entries()) {
    out.set(entry.path.replace(/^\/+/, ""), await entry.bytes());
  }
  return out;
}

/** A package whose body is the anchored shape. */
async function fixture(wsp = WSP): Promise<Uint8Array> {
  const files = await entries(await Io.toBuffer(Document.create()));
  files.set(
    "word/document.xml",
    encoder.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
        ANCHORED_SHAPE.replace(WSP, wsp) +
        `<w:sectPr/></w:body></w:document>`
    )
  );
  const archive = zip();
  for (const [path, data] of files) {
    archive.add(path, data);
  }
  return archive.toBuffer();
}

function findElement(el: XmlElement, name: string): XmlElement | undefined {
  if (el.name === name) {
    return el;
  }
  for (const child of el.children) {
    if (child.type === "element") {
      const found = findElement(child, name);
      if (found) {
        return found;
      }
    }
  }
  return undefined;
}

/**
 * Canonical form of an element: names, attributes sorted, children in order,
 * namespace declarations dropped. The drawing id is the writer's to assign,
 * and the text box's paragraphs are compared as text — the paragraph writer's
 * own form is not what this is about.
 */
function canonical(el: XmlElement): string {
  if (el.name === "w:txbxContent") {
    return `<w:txbxContent>${JSON.stringify(textOf(el))}</w:txbxContent>`;
  }
  const attrs = Object.entries(el.attributes)
    .filter(([k]) => !k.startsWith("xmlns") && !(el.name === "wps:cNvPr" && k === "id"))
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([k, v]) => ` ${k}="${v}"`)
    .join("");
  const children = el.children
    .filter((c): c is XmlElement => c.type === "element")
    .map(canonical)
    .join("");
  return `<${el.name}${attrs}>${children}</${el.name}>`;
}

function textOf(el: XmlElement): string {
  return el.children
    .map(c => (c.type === "text" ? c.value : c.type === "element" ? textOf(c) : ""))
    .join("");
}

async function writtenWsp(doc: DocxDocument): Promise<XmlElement> {
  const xml = decoder.decode((await entries(await Io.toBuffer(doc))).get("word/document.xml")!);
  return findElement(parsePartXml(xml, undefined).root, "wps:wsp")!;
}

const original = findElement(
  parsePartXml(`<w:root ${NS}>${WSP}</w:root>`, undefined).root,
  "wps:wsp"
)!;

describe("DrawingML shape fidelity", () => {
  it("reads the fixture as a modelled shape", async () => {
    const read = await Io.read(await fixture());
    const shape = read.body.find((b): b is DrawingShape => b.type === "drawingShape");
    expect(shape).toBeDefined();
    expect(shape!.textBodyAnchor).toBe("ctr");
    expect(shape!.outlineWidth).toBe(28575);
    expect(shape!.rotation).toBe(600000);
  });

  it("round-trips the wsp subtree unchanged", async () => {
    const read = await Io.read(await fixture());
    expect(canonical(await writtenWsp(read))).toBe(canonical(original));
    // And again, from what it wrote.
    const twice = await Io.read(await Io.toBuffer(read));
    expect(canonical(await writtenWsp(twice))).toBe(canonical(original));
  });

  it("an edited field replaces the markup it was read from", async () => {
    const read = await Io.read(await fixture());
    const index = read.body.findIndex(b => b.type === "drawingShape");
    const shape = read.body[index] as DrawingShape;
    const edited: DrawingShape = {
      ...shape,
      shapeType: "ellipse",
      fillColor: "00FF00",
      outlineColor: "123456",
      textBodyAnchor: "b"
    };
    const doc: DocxDocument = {
      ...read,
      body: read.body.map((b, i) => (i === index ? edited : b))
    };
    const wsp = canonical(await writtenWsp(doc));
    // Geometry, fill and line colour follow the fields…
    expect(wsp).toContain('<a:prstGeom prst="ellipse">');
    expect(wsp).not.toContain("<a:custGeom>");
    expect(wsp).toContain('<a:solidFill><a:srgbClr val="00FF00"></a:srgbClr></a:solidFill>');
    expect(wsp).not.toContain('val="accent1"><a:lumMod');
    expect(wsp).toContain(
      '<a:ln algn="ctr" cap="rnd" cmpd="dbl" w="28575"><a:solidFill><a:srgbClr val="123456">'
    );
    expect(wsp).not.toContain("<a:gradFill>");
    expect(wsp).toContain('anchor="b"');
    // …while what no field covers is kept.
    expect(wsp).toContain("<a:custDash>");
    expect(wsp).toContain('<a:tailEnd type="stealth">');
    expect(wsp).toContain("<wps:style>");
    expect(wsp).toContain('lIns="91440"');
    expect(wsp).toContain("<a:noAutofit>");
  });

  it("a text box holding a table keeps it, through a merge too", async () => {
    const table =
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const wsp = WSP.replace(
      "<w:p><w:r><w:t>Inside</w:t></w:r></w:p>",
      `<w:p><w:r><w:t>Inside</w:t></w:r></w:p>${table}`
    );
    const read = await Io.read(await fixture(wsp));
    const xml = (doc: DocxDocument) =>
      writtenWsp(doc).then(el => {
        const content = findElement(el, "w:txbxContent")!;
        return content.children
          .filter((c): c is XmlElement => c.type === "element")
          .map(c => c.name);
      });
    expect(await xml(read)).toEqual(["w:p", "w:tbl"]);
    expect(await xml(Transform.merge([Document.create(), read]))).toEqual(["w:p", "w:tbl"]);

    // Editing the text replaces the content as read.
    const index = read.body.findIndex(b => b.type === "drawingShape");
    const shape = read.body[index] as DrawingShape;
    const edited: DrawingShape = { ...shape, textContent: [...shape.textContent!] };
    expect(
      await xml({ ...read, body: read.body.map((b, i) => (i === index ? edited : b)) })
    ).toEqual(["w:p"]);
  });

  it("a shape without a transform is written without one", async () => {
    const wsp = WSP.replace(/<a:xfrm rot="600000">.*?<\/a:xfrm>/, "");
    const read = await Io.read(await fixture(wsp));
    const expected = findElement(
      parsePartXml(`<w:root ${NS}>${wsp}</w:root>`, undefined).root,
      "wps:wsp"
    )!;
    expect(canonical(await writtenWsp(read))).toBe(canonical(expected));
  });

  it("the raw-XML policy governs the kept markup", async () => {
    const read = await Io.read(await fixture());
    await expect(
      Io.toBuffer(read, { securityPolicy: { rawXmlPolicy: "reject" } })
    ).rejects.toMatchObject({ name: "DocxRawXmlPolicyError", site: "drawingShape._preserved" });
    const stripped = await Io.toBuffer(read, { securityPolicy: { rawXmlPolicy: "strip" } });
    const xml = decoder.decode((await entries(stripped)).get("word/document.xml")!);
    expect(xml).toContain("<wps:wsp");
    expect(xml).not.toContain("<wps:style>");
    expect(xml).not.toContain("<a:custGeom>");
  });
});
