/**
 * A subtree the reader keeps as XML and writes back verbatim has left its ancestors behind. It must still
 * mean, inside the part it is written into, exactly what it meant inside the source: every prefix it names
 * declared, and every unprefixed name in the namespace it was in.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { expandedNames } from "@excel/__tests__/helpers/expanded-names";
import { Cell, Workbook } from "@excel/index";
import { ChartSpaceXform } from "@excel/xlsx/xform/chart/chart-space-xform";
import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { createPrefixNormalizer } from "@excel/xlsx/xform/namespace-prefixes";
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { TableXform } from "@excel/xlsx/xform/table/table-xform";
import { describe, expect, it } from "vitest";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/** `xml` read by `xform` and written back by it: the part as the writer emits it. */
async function roundTrip(
  xform: { toXml(model: unknown): string; prepare(model: unknown, options: object): void },
  xml: string
): Promise<string> {
  const model = await parseXformStream(xform as never, [xml]);
  if (xform instanceof TableXform) {
    xform.prepare(model as never, {});
  }
  return xform.toXml(model);
}

/** The subtree named `within` means the same in `written` as in `source`, and names no unbound prefix. */
function expectSameSubtree(source: string, written: string, within: string): void {
  const expected = expandedNames(source, within);
  expect(expected.length).toBeGreaterThan(1);
  expect(expandedNames(written, within)).toEqual(expected);
}

const GROUP = (inner = "") =>
  `<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
  `<xdr:grpSp${inner}><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Group 9"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
  `<xdr:grpSpPr/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="11" name="Inner"><a:hlinkClick r:id="rId7"/></xdr:cNvPr>` +
  `<xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId5"/></xdr:blipFill><xdr:spPr/></xdr:pic>` +
  `</xdr:grpSp><xdr:clientData/>`;

describe("a captured subtree declares the prefixes it borrowed from its ancestors", () => {
  it("from the document element, when the writer does not write that declaration back", async () => {
    // The drawing writer declares `xdr` and `a` only; `r` was the source root's.
    const source =
      `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${REL}">` +
      `<xdr:twoCellAnchor>${GROUP()}</xdr:twoCellAnchor></xdr:wsDr>`;
    const written = await roundTrip(new DrawingXform(), source);
    expect(written).toContain(`<xdr:grpSp xmlns:r="${REL}">`);
    expectSameSubtree(source, written, `{${XDR}}grpSp`);
  });

  it("from an ancestor the reader models, whose declarations it does not keep", async () => {
    const source =
      `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">` +
      `<xdr:twoCellAnchor xmlns:r="${REL}">${GROUP()}</xdr:twoCellAnchor></xdr:wsDr>`;
    const written = await roundTrip(new DrawingXform(), source);
    expectSameSubtree(source, written, `{${XDR}}grpSp`);
  });

  it("but not one the writer declares anyway, so Excel's own parts keep their bytes", async () => {
    const source =
      `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">` +
      `<xdr:twoCellAnchor>${GROUP(` xmlns:r="${REL}"`)}</xdr:twoCellAnchor></xdr:wsDr>`;
    const written = await roundTrip(new DrawingXform(), source);
    expect(written).toContain(`<xdr:grpSp xmlns:r="${REL}">`);
    expect(written.match(/xmlns:a=/g)).toHaveLength(1);
    expectSameSubtree(source, written, `{${XDR}}grpSp`);
  });

  it("in a chart, from a structural ancestor", async () => {
    const source =
      `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${REL}"><c:chart>` +
      `<c:plotArea xmlns:q="urn:q"><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      `<c:spPr><a:ln q:style="x"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>` +
      `</c:plotArea></c:chart></c:chartSpace>`;
    const written = await roundTrip(new ChartSpaceXform(), source);
    expect(written).toContain(`<c:spPr xmlns:q="urn:q">`);
    expectSameSubtree(source, written, `{${C}}spPr`);
  });

  it("but not one an enclosing element it keeps writes back itself", async () => {
    const source =
      `<table xmlns="${MAIN}" id="1" name="T" displayName="T" ref="A1:A3">` +
      `<autoFilter ref="A1:A3" xmlns:q="urn:q"><filterColumn colId="0">` +
      `<customFilters><customFilter q:x="1" val="5"/></customFilters></filterColumn></autoFilter>` +
      `<tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns>` +
      `<tableStyleInfo name="TableStyleMedium2"/></table>`;
    const written = await roundTrip(new TableXform(), source);
    expect(written).toContain(`<customFilters><customFilter q:x="1" val="5"/></customFilters>`);
    expectSameSubtree(source, written, `{${MAIN}}customFilters`);
  });
});

describe("a captured subtree keeps its unprefixed names in their namespaces", () => {
  it("where a prefixed namespace became the default one, an element in no namespace stays in none", async () => {
    const source =
      `<t:table xmlns:t="${MAIN}" id="1" name="T" displayName="T" ref="A1:A3">` +
      `<t:autoFilter ref="A1:A3"><t:filterColumn colId="0"><t:filters><t:filter val="1"/>` +
      `<foo><t:bar/></foo></t:filters></t:filterColumn></t:autoFilter>` +
      `<t:tableColumns count="1"><t:tableColumn id="1" name="A"/></t:tableColumns>` +
      `<t:tableStyleInfo name="TableStyleMedium2"/></t:table>`;
    const written = await roundTrip(new TableXform(), source);
    expect(written).toContain(`<foo xmlns=""><bar xmlns="${MAIN}"/></foo>`);
    expectSameSubtree(source, written, `{${MAIN}}filters`);
  });

  it("through a worksheet read and written as a workbook", async () => {
    const wb = Workbook.create();
    Cell.setValue(Workbook.addWorksheet(wb, "S"), "A1", 1);
    const entries = await extractAll(await Workbook.toBuffer(wb));
    const sheet =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<s:worksheet xmlns:s="${MAIN}" xmlns:r="${REL}"><s:sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c>` +
      `</s:row></s:sheetData><s:autoFilter ref="A1:A2"><s:filterColumn colId="0"><s:filters>` +
      `<s:filter val="1"/><foo/></s:filters></s:filterColumn></s:autoFilter></s:worksheet>`;
    const edited = createZip(
      [...entries].map(([name, file]) => ({
        name,
        data: name === "xl/worksheets/sheet1.xml" ? new TextEncoder().encode(sheet) : file.data
      }))
    );
    const back = Workbook.create();
    await Workbook.read(back, await edited);
    const written = new TextDecoder().decode(
      (await extractAll(await Workbook.toBuffer(back))).get("xl/worksheets/sheet1.xml")!.data
    );
    expect(written).toContain(`<foo xmlns=""/>`);
    expectSameSubtree(sheet, written, `{${MAIN}}filterColumn`);
  });
});

describe("the lazy normaliser tracks what a capture needs", () => {
  it("knows the document element's declarations while dormant", () => {
    const prefixes = createPrefixNormalizer({ lazy: true });
    expect(
      prefixes.scan(
        `<?xml version="1.0"?><!-- <x:y xmlns:x="u"> --><worksheet xmlns="${MAIN}" xmlns:r="${REL}">`
      )
    ).toBe(false);
    expect(prefixes.scan(`<sheetData/>`)).toBe(false);
    expect(Object.fromEntries(prefixes.bindings())).toEqual({ "": MAIN, r: REL });
    expect(Object.fromEntries(prefixes.documentBindings())).toEqual({ "": MAIN, r: REL });
  });

  it("wakes for a declaration below the document element, canonical or not", () => {
    const prefixes = createPrefixNormalizer({ lazy: true });
    prefixes.scan(`<worksheet xmlns="${MAIN}"><sheetData/>`);
    expect(
      prefixes.scan(
        `<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">`
      )
    ).toBe(true);
  });

  it("wakes when the document element's start tag is cut by a chunk", () => {
    const prefixes = createPrefixNormalizer({ lazy: true });
    expect(prefixes.scan(`<worksheet xmlns="${MAIN}" xmlns:s="${MAIN}"`)).toBe(false);
    expect(prefixes.scan(`>`)).toBe(true);
  });
});
