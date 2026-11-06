/**
 * An anchor's `<mc:AlternateContent>` names the namespaces its Choice requires by **prefix**, so the
 * writer must declare every prefix it writes there, for the namespace the source meant — and keep the
 * Fallback the source gave consumers that cannot take the Choice.
 */
import { expandedNames } from "@excel/__tests__/helpers/expanded-names";
import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { describe, expect, it } from "vitest";

const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const CX2 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";
const A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

const FROM_TO =
  `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;

const CHART_EX_FRAME =
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/>` +
  `</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
  `<a:graphicData uri="${CX}"><cx:chart xmlns:cx="${CX}" r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame>`;

const OWN_FALLBACK =
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/>` +
  `<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>The source placeholder</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;

function drawing(anchors: string, prefixes = `xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}"`) {
  return `<xdr:wsDr ${prefixes}>${anchors}</xdr:wsDr>`;
}

function chartExAnchor(choice: string, fallback = OWN_FALLBACK, mc = "mc") {
  return (
    `<xdr:twoCellAnchor>${FROM_TO}<${mc}:AlternateContent xmlns:${mc}="${MC}">` +
    `<${mc}:Choice ${choice}>${CHART_EX_FRAME}</${mc}:Choice><${mc}:Fallback>${fallback}</${mc}:Fallback>` +
    `</${mc}:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
  );
}

async function roundTrip(xml: string): Promise<string> {
  const xform = new DrawingXform();
  return xform.toXml(await parseXformStream(xform, [xml]));
}

/** Every `mc:Choice` start tag in `xml`. */
function choices(xml: string): string[] {
  return xml.match(/<mc:Choice[^>]*>/g) ?? [];
}

describe("an anchor's AlternateContent is written with the namespaces it requires declared", () => {
  it("keeps a ChartEx Choice that requires another revision in the anchor, declared", async () => {
    const written = await roundTrip(drawing(chartExAnchor(`xmlns:cx2="${CX2}" Requires="cx2"`)));
    // In the shape slot, not around the anchor, with the prefix it names declared.
    expect(written).toMatch(
      /<xdr:twoCellAnchor>(?:(?!<\/xdr:twoCellAnchor>).)*<mc:AlternateContent/s
    );
    expect(choices(written)).toEqual([`<mc:Choice xmlns:cx2="${CX2}" Requires="cx2">`]);
    // Every prefix the part names — `Requires` included — is declared.
    expect(() => expandedNames(written)).not.toThrow();
  });

  it("reads a ChartEx revision bound to another prefix by its namespace", async () => {
    const written = await roundTrip(drawing(chartExAnchor(`xmlns:zcx1="${CX1}" Requires="zcx1"`)));
    expect(choices(written)).toEqual([`<mc:Choice xmlns:cx1="${CX1}" Requires="cx1">`]);
    expect(() => expandedNames(written)).not.toThrow();
  });

  it("resolves a Choice prefix declared on an ancestor", async () => {
    const written = await roundTrip(
      drawing(
        chartExAnchor(`Requires="cx2"`),
        `xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}" xmlns:cx2="${CX2}"`
      )
    );
    expect(choices(written)).toEqual([`<mc:Choice xmlns:cx2="${CX2}" Requires="cx2">`]);
  });

  it("writes back the source's Fallback, and Excel's placeholder only where there was none", async () => {
    const own = await roundTrip(drawing(chartExAnchor(`xmlns:cx1="${CX1}" Requires="cx1"`)));
    expect(own).toContain("The source placeholder");
    expect(own).not.toContain("This chart isn");
    const none = await roundTrip(drawing(chartExAnchor(`xmlns:cx1="${CX1}" Requires="cx1"`, "")));
    expect(none).toContain("This chart isn");
  });

  it("keeps a form control's Choice around the anchor, declared, with its Fallback", async () => {
    const control =
      `<xdr:twoCellAnchor editAs="absolute">${FROM_TO}<xdr:sp macro="" textlink=""><xdr:nvSpPr>` +
      `<xdr:cNvPr id="1025" name="Check Box 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>` +
      `<xdr:clientData/></xdr:twoCellAnchor>`;
    const fallbackAnchor = control.replace("Check Box 1", "Fallback Box");
    const source = drawing(
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${control}` +
        `</mc:Choice><mc:Fallback>${fallbackAnchor}</mc:Fallback></mc:AlternateContent>`
    );
    const written = await roundTrip(source);
    expect(choices(written)).toEqual([`<mc:Choice xmlns:a14="${A14}" Requires="a14">`]);
    expect(written).toMatch(/<mc:AlternateContent[^>]*><mc:Choice[^>]*><xdr:twoCellAnchor/);
    expect(written).toContain("Fallback Box");
    expect(() => expandedNames(written)).not.toThrow();
  });

  it("writes no Choice for a prefix no namespace can be declared for", async () => {
    const xform = new DrawingXform();
    const model = await parseXformStream(xform, [
      drawing(
        `<xdr:twoCellAnchor>${FROM_TO}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="S"/>` +
          `<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`
      )
    ]);
    (model as { anchors: { alternateContent?: unknown }[] }).anchors[0].alternateContent = {
      requires: "unbound"
    };
    const written = xform.toXml(model);
    expect(written).not.toContain("mc:Choice");
    expect(() => expandedNames(written)).not.toThrow();
  });
});

const PICTURE =
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="5" name="Picture 5"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
  `<xdr:blipFill><a:blip r:embed="rId5"/></xdr:blipFill><xdr:spPr/></xdr:pic>`;

describe("an anchor wrapped in an AlternateContent that holds another keeps both", () => {
  const inner = (choice: string, content: string, fallback: string) =>
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice ${choice}>${content}</mc:Choice>` +
    `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const shape = (name: string) =>
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="4" name="${name}"/><xdr:cNvSpPr/>` +
    `</xdr:nvSpPr><xdr:spPr/></xdr:sp>`;
  const outer = (anchor: string) =>
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${anchor}` +
    `</mc:Choice><mc:Fallback>${shapeAnchor("Outer fallback")}</mc:Fallback></mc:AlternateContent>`;
  const shapeAnchor = (name: string) =>
    `<xdr:twoCellAnchor editAs="absolute">${FROM_TO}${shape(name)}<xdr:clientData/></xdr:twoCellAnchor>`;

  for (const [label, content] of [
    ["a ChartEx frame", inner(`xmlns:cx2="${CX2}" Requires="cx2"`, CHART_EX_FRAME, OWN_FALLBACK)],
    ["a picture", inner(`xmlns:a14="${A14}" Requires="a14"`, PICTURE, shape("Inner fallback"))]
  ] as const) {
    it(`around ${label}`, async () => {
      const source = drawing(
        outer(
          `<xdr:twoCellAnchor editAs="absolute">${FROM_TO}${content}<xdr:clientData/></xdr:twoCellAnchor>`
        )
      );
      const written = await roundTrip(source);
      // Outer around the anchor, inner in its shape slot — both declared, both Fallbacks kept.
      expect(written).toMatch(
        /<mc:AlternateContent[^>]*><mc:Choice[^>]*Requires="a14"><xdr:twoCellAnchor[^>]*><xdr:from>.*<\/xdr:to><mc:AlternateContent/s
      );
      expect(choices(written)).toEqual([
        `<mc:Choice xmlns:a14="${A14}" Requires="a14">`,
        label === "a picture"
          ? `<mc:Choice xmlns:a14="${A14}" Requires="a14">`
          : `<mc:Choice xmlns:cx2="${CX2}" Requires="cx2">`
      ]);
      expect(written).toContain("Outer fallback");
      expect(written).toContain(
        label === "a picture" ? "Inner fallback" : "The source placeholder"
      );
      expect(() => expandedNames(written)).not.toThrow();
      // And it reads back the same.
      expect(await roundTrip(written)).toBe(written);
    });
  }
});
