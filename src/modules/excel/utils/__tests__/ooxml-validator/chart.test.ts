/**
 * Chart & ChartEx internal-structure checks.
 *
 * Negative samples for every `chartEx-*` problem kind the validator
 * emits. Each test seeds a minimal fixture with the exact broken
 * pattern the writer used to emit, then asserts the validator produces
 * the matching `kind`. Positive baseline tests cover the "schema
 * conformant" shape so we do not accidentally flag valid output.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith } from "./fixtures";

// -----------------------------------------------------------------------------
// ChartEx fixture helpers
// -----------------------------------------------------------------------------

/**
 * Wrap a chartEx `<cx:chartSpace>` body in the minimum scaffolding
 * needed for the validator to find it — the part itself, the override
 * content-type, and a drawing rel that points at it (so the validator
 * doesn't fail on a missing parent anchor first).
 */
function buildChartExPackage(chartExInner: string): Uint8Array {
  const parts = baseParts();

  // Add a drawing that hosts the chartEx.
  parts["xl/drawings/drawing1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame macro="">
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback><xdr:sp/></mc:Fallback>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
  parts["xl/charts/chartEx1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
${chartExInner}
</cx:chartSpace>`;

  // Content types and rels.
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
    { partName: "/xl/charts/chartEx1.xml", contentType: "application/vnd.ms-office.chartex+xml" },
    {
      partName: "/xl/drawings/drawing1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
    }
  ]);
  return buildPackage(parts);
}

function minimalChartExBody(): string {
  // A valid baseline chartEx with one data entry, one axis, one series.
  // ChartEx requires `<cx:f>` to point at a hidden defined name, not
  // directly at a worksheet range, so the baseline uses the
  // `_xlchart.v1.N` convention Excel itself writes.
  return `
  <cx:chartData>
    <cx:data id="0">
      <cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`;
}

// -----------------------------------------------------------------------------
// Tests
// -----------------------------------------------------------------------------

describe("ooxml-validator / chart — baseline", () => {
  it("accepts a minimal schema-valid chartEx", async () => {
    const report = await validateXlsxBuffer(buildChartExPackage(minimalChartExBody()));
    const chartProblems = report.problems.filter(p => p.kind.startsWith("chart"));
    expect(chartProblems).toEqual([]);
  });
});

describe("ooxml-validator / chart — chartEx schema violations", () => {
  it("flags a series with more than one <cx:dataId>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim></cx:data>
    <cx:data id="1"><cx:numDim type="val"><cx:f>_xlchart.v1.1</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:dataId val="1"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-series-too-many-dataId")).toBe(true);
  });

  it("flags <cx:axisId>N</cx:axisId> text-content form", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:axisId>2</cx:axisId>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-typed-element-text-form")).toBe(true);
  });

  it("flags <cx:binCount>N</cx:binCount> text-content form", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:binning><cx:binCount>12</cx:binCount></cx:binning>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-typed-element-text-form")).toBe(true);
  });

  it("flags <cx:auto/> element", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:binning><cx:auto/></cx:binning>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-invalid-auto-element")).toBe(true);
  });

  it("flags <cx:paretoLine> inside <cx:layoutPr>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:paretoLine/>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-paretoLine-in-layoutPr")).toBe(true);
  });

  it("flags <cx:title><cx:layout/>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:title pos="t" align="ctr" overlay="0"><cx:layout/></cx:title>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-title-direct-layout")).toBe(true);
  });
});

describe("ooxml-validator / chart — Tier 2 semantic checks", () => {
  it("flags axis missing both pos and type", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
        <cx:axis id="0"/>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-axis-missing-pos-and-type")).toBe(true);
  });

  it("accepts axis with pos only", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
        <cx:axis id="0" pos="b"/>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-axis-missing-pos-and-type")).toBe(false);
  });

  it("flags <cx:f> pointing at a direct sheet range", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$1:$A$3</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(
      report.problems.some(p => p.kind === "chartEx-f-uses-direct-range-not-defined-name")
    ).toBe(true);
  });

  it("accepts <cx:f> pointing at a hidden defined name (_xlchart.v1.N)", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(
      report.problems.some(p => p.kind === "chartEx-f-uses-direct-range-not-defined-name")
    ).toBe(false);
  });

  it("flags waterfall series missing <cx:layoutPr><cx:subtotals>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="waterfall" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-waterfall-missing-subtotals")).toBe(true);
  });

  it("accepts waterfall series with empty <cx:subtotals/> marker", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="waterfall" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr><cx:subtotals/></cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-waterfall-missing-subtotals")).toBe(false);
  });
});
