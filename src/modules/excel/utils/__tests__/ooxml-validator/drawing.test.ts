/**
 * Drawing part — chartEx AlternateContent wrapping invariants.
 *
 * Three `drawing-chartEx-*` problem kinds mirror the three "Removed
 * Part: /xl/drawings/drawingN.xml (Drawing shape)" patterns that real
 * Excel logs attribute to broken chartEx drawings:
 *
 *   - `<cx:chart>` reference without any `<mc:AlternateContent>` wrap.
 *   - `<mc:AlternateContent>` wraps the ENTIRE anchor (duplicates from/to).
 *   - `<mc:Fallback>` is missing or empty.
 *
 * The positive baseline at the top is the shape Excel itself writes.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith } from "./fixtures";

function buildDrawingPackage(drawingXml: string): Uint8Array {
  const parts = baseParts();
  parts["xl/drawings/drawing1.xml"] = drawingXml;
  parts["xl/charts/chartEx1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$1:$A$2</cx:f></cx:numDim></cx:data></cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0"><cx:dataId val="0"/></cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;
  parts["[Content_Types].xml"] = contentTypesWith([
    {
      partName: "/xl/workbook.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    },
    {
      partName: "/xl/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    },
    {
      partName: "/xl/sharedStrings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    },
    {
      partName: "/xl/theme/theme1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.theme+xml"
    },
    {
      partName: "/xl/worksheets/sheet1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    },
    {
      partName: "/xl/drawings/drawing1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
    },
    { partName: "/xl/charts/chartEx1.xml", contentType: "application/vnd.ms-office.chartex+xml" }
  ]);
  return buildPackage(parts);
}

describe("ooxml-validator / drawing — chartEx AlternateContent wrapping", () => {
  it("accepts the canonical Microsoft-Excel layout", async () => {
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback><xdr:sp/></mc:Fallback>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`)
    );
    const drawingProblems = report.problems.filter(p => p.kind.startsWith("drawing-chartEx"));
    expect(drawingProblems).toEqual([]);
  });

  it("flags a bare <xdr:graphicFrame><cx:chart/> without AlternateContent", async () => {
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`)
    );
    expect(
      report.problems.some(p => p.kind === "drawing-chartEx-missing-alternateContent-wrap")
    ).toBe(true);
  });

  it("flags an empty <mc:Fallback>", async () => {
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback/>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`)
    );
    expect(
      report.problems.some(p => p.kind === "drawing-chartEx-alternateContent-empty-fallback")
    ).toBe(true);
  });

  it("flags AlternateContent that wraps the entire anchor", async () => {
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:AlternateContent>
    <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
      <xdr:twoCellAnchor>
        <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
        <xdr:clientData/>
      </xdr:twoCellAnchor>
    </mc:Choice>
    <mc:Fallback><xdr:sp/></mc:Fallback>
  </mc:AlternateContent>
</xdr:wsDr>`)
    );
    expect(
      report.problems.some(p => p.kind === "drawing-chartEx-alternateContent-outer-wrap")
    ).toBe(true);
  });

  it("flags chartEx drawing missing a16:creationId extension (warning)", async () => {
    // includeWarnings must be true for the warning to surface.
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback><xdr:sp/></mc:Fallback>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "drawing-chartEx-missing-creationId" && p.severity === "warning"
      )
    ).toBe(true);
  });

  it("accepts chartEx drawing WITH a16:creationId extension", async () => {
    const report = await validateXlsxBuffer(
      buildDrawingPackage(`<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame>
          <xdr:nvGraphicFramePr>
            <xdr:cNvPr id="2" name="Chart 1">
              <a:extLst>
                <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                  <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0000-0000-0000-000000000000}"/>
                </a:ext>
              </a:extLst>
            </xdr:cNvPr>
            <xdr:cNvGraphicFramePr/>
          </xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback><xdr:sp/></mc:Fallback>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "drawing-chartEx-missing-creationId")).toBe(false);
  });
});
